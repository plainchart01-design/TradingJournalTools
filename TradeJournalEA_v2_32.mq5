//+------------------------------------------------------------------+
//|  TradeJournalEA.mq5  (v2.32)                                     |
//|  Sends every trade open/close to your Google Sheets journal      |
//|                                                                  |
//|  SETUP:                                                          |
//|  1. Paste your Apps Script deployment URL into WEBHOOK_URL below |
//|  2. In MetaEditor press F7 to compile                           |
//|  3. Attach to ANY or ALL charts — safe to run on multiple pairs  |
//|  4. MT5 -> Tools -> Options -> Expert Advisors:                  |
//|       Allow WebRequest -> add: https://script.google.com         |
//|                                                                  |
//|  WHAT'S NEW IN v2.32:                                            |
//|  - Post() now recognizes "status":"updated" (a successful close  |
//|    that updated an existing row) and "status":"duplicate" (a     |
//|    deliberate skip) as their own cases instead of lumping them   |
//|    into the "⚠ Unexpected response" warning. Closes were         |
//|    actually saving fine after the v2.31/v3 backend update — this |
//|    was a misleading log message, not a data problem.             |
//|                                                                  |
//|  WHAT'S NEW IN v2.31:                                            |
//|  - Open and close reports now include the position ticket. The   |
//|    backend uses it to tell "two trades opened in the same        |
//|    minute" apart, and to UPDATE the original open row when a     |
//|    trade closes instead of bouncing it as a duplicate. Requires  |
//|    the matching Apps Script v3 update — old backend code will    |
//|    just ignore the extra "ticket" field harmlessly.              |
//|                                                                  |
//|  WHAT'S NEW IN v2.30:                                            |
//|  - Uses GlobalVariableSetOnCondition() — an atomic compare-and- |
//|    swap — to guarantee only ONE instance sends each trade even   |
//|    when the EA is attached to multiple charts simultaneously.    |
//|    Safe to put on every chart you trade.                         |
//+------------------------------------------------------------------+
#property copyright "TradeJournal"
#property version   "2.32"
#property strict

input string WEBHOOK_URL = "PASTE_YOUR_APPS_SCRIPT_URL_HERE";

#define GV_OPEN  "TJ_OPEN_"
#define GV_CLOSE "TJ_CLOSE_"

// Live SL/TP tracking per instance
ulong  trkTicket[];
double trkSL[];
double trkTP[];

//+------------------------------------------------------------------+
int OnInit()
  {
   ArrayResize(trkTicket, 0);
   ArrayResize(trkSL,     0);
   ArrayResize(trkTP,     0);
   Print("TradeJournalEA v2.32: Started on ", Symbol(), ". Safe to run on multiple charts.");
   return(INIT_SUCCEEDED);
  }

//+------------------------------------------------------------------+
void OnTick()
  {
   UpdateTracking();
   CheckNewPositions();
   CheckClosedDeals();
  }

// ── ATOMIC CLAIM — the core of duplicate prevention ──────────────
//
// GlobalVariableSetOnCondition(name, new, expected) is an MT5 atomic
// compare-and-swap. It sets the variable to `new` ONLY IF the current
// value equals `expected`, and returns true if it succeeded.
//
// Two EA instances running simultaneously both call this for the same
// ticket. Only ONE will see the value at 0.0 and flip it to 1.0 —
// the other arrives a microsecond later, sees 1.0, and returns false.
// That loser skips the send entirely. No duplicates possible.
//
// Returns true  → this instance won the claim, should send.
// Returns false → another instance already claimed it, skip.
// ─────────────────────────────────────────────────────────────────
bool ClaimOpen(ulong ticket)
  {
   string name = GV_OPEN + IntegerToString((long)ticket);
   // Ensure the variable exists before the atomic swap
   if(!GlobalVariableCheck(name)) GlobalVariableSet(name, 0.0);
   return GlobalVariableSetOnCondition(name, 1.0, 0.0);
  }

bool ClaimClose(ulong deal)
  {
   string name = GV_CLOSE + IntegerToString((long)deal);
   if(!GlobalVariableCheck(name)) GlobalVariableSet(name, 0.0);
   return GlobalVariableSetOnCondition(name, 1.0, 0.0);
  }

// ── LIVE SL/TP TRACKING ──────────────────────────────────────────
void UpdateTracking()
  {
   for(int i = 0; i < PositionsTotal(); i++)
     {
      ulong ticket = PositionGetTicket(i);
      if(ticket == 0 || !PositionSelectByTicket(ticket)) continue;
      double sl = PositionGetDouble(POSITION_SL);
      double tp = PositionGetDouble(POSITION_TP);
      int idx = TrackedIndex(ticket);
      if(idx == -1)
        {
         int n = ArraySize(trkTicket);
         ArrayResize(trkTicket,n+1); ArrayResize(trkSL,n+1); ArrayResize(trkTP,n+1);
         trkTicket[n]=ticket; trkSL[n]=sl; trkTP[n]=tp;
        }
      else { trkSL[idx]=sl; trkTP[idx]=tp; }
     }
  }

int TrackedIndex(ulong ticket)
  { for(int i=0;i<ArraySize(trkTicket);i++) if(trkTicket[i]==ticket) return i; return -1; }

void RemoveTracked(ulong ticket)
  {
   int n=ArraySize(trkTicket);
   for(int i=0;i<n;i++)
     {
      if(trkTicket[i]!=ticket) continue;
      for(int j=i;j<n-1;j++){trkTicket[j]=trkTicket[j+1];trkSL[j]=trkSL[j+1];trkTP[j]=trkTP[j+1];}
      ArrayResize(trkTicket,n-1); ArrayResize(trkSL,n-1); ArrayResize(trkTP,n-1);
      return;
     }
  }

// ── CHECK NEW (OPEN) POSITIONS ────────────────────────────────────
void CheckNewPositions()
  {
   for(int i = 0; i < PositionsTotal(); i++)
     {
      ulong ticket = PositionGetTicket(i);
      if(ticket == 0) continue;
      if(!ClaimOpen(ticket)) continue;  // atomic — only ONE instance passes
      if(!PositionSelectByTicket(ticket)) continue;

      string pair  = PositionGetString(POSITION_SYMBOL);
      string dir   = (PositionGetInteger(POSITION_TYPE)==POSITION_TYPE_BUY)?"LONG":"SHORT";
      double entry = PositionGetDouble(POSITION_PRICE_OPEN);
      double sl    = PositionGetDouble(POSITION_SL);
      double tp    = PositionGetDouble(POSITION_TP);
      datetime dt  = (datetime)PositionGetInteger(POSITION_TIME);
      int digits   = (int)SymbolInfoInteger(pair,SYMBOL_DIGITS);

      string json="{";
      json+="\"ticket\":"      +Q(IntegerToString((long)ticket)) +",";
      json+="\"pair\":"        +Q(pair)              +",";
      json+="\"direction\":"   +Q(dir)               +",";
      json+="\"entry\":"       +Q(Dbl(entry,digits)) +",";
      json+="\"stop_loss\":"   +Q(Dbl(sl,digits))    +",";
      json+="\"take_profit\":" +Q(Dbl(tp,digits))    +",";
      json+="\"time_open\":"   +Q(Fmt(dt))           +",";
      json+="\"day\":"         +Q(DoW(dt))           +",";
      json+="\"session\":"     +Q(Sess(dt))          +",";
      json+="\"outcome\":"     +Q("OPEN")            +",";
      json+="\"rr\":"          +Q("")                +",";
      json+="\"entry_type\":"  +Q("")                +",";
      json+="\"emotion\":"     +Q("")                +",";
      json+="\"source\":"      +Q("MetaTrader5");
      json+="}";
      Post(json);
     }
  }

// ── CHECK CLOSED DEALS ────────────────────────────────────────────
void CheckClosedDeals()
  {
   if(!HistorySelect(TimeCurrent()-86400,TimeCurrent())) return;

   for(int i = 0; i < HistoryDealsTotal(); i++)
     {
      ulong deal = HistoryDealGetTicket(i);
      if(deal == 0) continue;
      if(HistoryDealGetInteger(deal,DEAL_ENTRY)!=DEAL_ENTRY_OUT) continue;
      if(!ClaimClose(deal)) continue;   // atomic — only ONE instance passes

      ulong posId=(ulong)HistoryDealGetInteger(deal,DEAL_POSITION_ID);

      string pair    = HistoryDealGetString(deal,DEAL_SYMBOL);
      double profit  = HistoryDealGetDouble(deal,DEAL_PROFIT);
      double closeP  = HistoryDealGetDouble(deal,DEAL_PRICE);
      datetime closeDt=(datetime)HistoryDealGetInteger(deal,DEAL_TIME);

      double openP=0,sl=0,tp=0;
      datetime openDt=0;
      string dir="LONG";

      for(int j=0;j<HistoryDealsTotal();j++)
        {
         ulong d2=HistoryDealGetTicket(j);
         if((ulong)HistoryDealGetInteger(d2,DEAL_POSITION_ID)!=posId) continue;
         if(HistoryDealGetInteger(d2,DEAL_ENTRY)!=DEAL_ENTRY_IN) continue;
         openP  = HistoryDealGetDouble(d2,DEAL_PRICE);
         sl     = HistoryDealGetDouble(d2,DEAL_SL);
         tp     = HistoryDealGetDouble(d2,DEAL_TP);
         openDt = (datetime)HistoryDealGetInteger(d2,DEAL_TIME);
         dir    = (HistoryDealGetInteger(d2,DEAL_TYPE)==DEAL_TYPE_BUY)?"LONG":"SHORT";
         break;
        }

      // Use live-tracked SL/TP if available
      int tIdx=TrackedIndex(posId);
      if(tIdx!=-1)
        {
         if(trkSL[tIdx]!=0.0) sl=trkSL[tIdx];
         if(trkTP[tIdx]!=0.0) tp=trkTP[tIdx];
         RemoveTracked(posId);
        }

      int    digits =(int)SymbolInfoInteger(pair,SYMBOL_DIGITS);
      string outcome=(profit>=0)?"WIN":"LOSS";

      string json="{";
      json+="\"ticket\":"      +Q(IntegerToString((long)posId))   +",";
      json+="\"pair\":"        +Q(pair)               +",";
      json+="\"direction\":"   +Q(dir)                +",";
      json+="\"entry\":"       +Q(Dbl(openP,digits))  +",";
      json+="\"stop_loss\":"   +Q(Dbl(sl,digits))     +",";
      json+="\"take_profit\":" +Q(Dbl(tp,digits))     +",";
      json+="\"time_open\":"   +Q(Fmt(openDt))        +",";
      json+="\"time_close\":"  +Q(Fmt(closeDt))       +",";
      json+="\"day\":"         +Q(DoW(openDt))        +",";
      json+="\"session\":"     +Q(Sess(openDt))       +",";
      json+="\"duration\":"    +Q(Dur(openDt,closeDt))+",";
      json+="\"outcome\":"     +Q(outcome)            +",";
      json+="\"rr\":"          +Q(RR(openP,sl,closeP))+",";
      json+="\"entry_type\":"  +Q("")                 +",";
      json+="\"emotion\":"     +Q("")                 +",";
      json+="\"source\":"      +Q("MetaTrader5");
      json+="}";
      Post(json);
     }
  }

// ── HTTP POST ─────────────────────────────────────────────────────
void Post(string json)
  {
   Print("TradeJournalEA: Sending -> ", json);
   char body[], resp[];
   string respHdr;
   StringToCharArray(json, body, 0, StringLen(json));
   int code = WebRequest("POST", WEBHOOK_URL, "Content-Type: application/json\r\n",
                          10000, body, resp, respHdr);
   if(code == -1)
     {
      Print("TradeJournalEA: ❌ WebRequest failed. Error=", GetLastError(),
            " — Check Tools→Options→Expert Advisors and make sure https://script.google.com is whitelisted.");
      return;
     }

   string respStr = CharArrayToString(resp);
   Print("TradeJournalEA: HTTP ", code, " -> ", respStr);

   // Parse the response to give a clear pass/fail message in the Experts log.
   // HTTP 200 only means Apps Script ran — we must check the JSON body for
   // "status":"success" to know the trade was actually written to the sheet.
   if(StringFind(respStr, "\"status\":\"success\"") >= 0 ||
      StringFind(respStr, "\"status\": \"success\"") >= 0)
     {
      Print("TradeJournalEA: ✅ Trade written to Google Sheet successfully.");
     }
   else if(StringFind(respStr, "\"status\":\"updated\"") >= 0 ||
           StringFind(respStr, "\"status\": \"updated\"") >= 0)
     {
      Print("TradeJournalEA: ✅ Trade closed — existing row updated in Google Sheet.");
     }
   else if(StringFind(respStr, "\"status\":\"duplicate\"") >= 0 ||
           StringFind(respStr, "\"status\": \"duplicate\"") >= 0)
     {
      Print("TradeJournalEA: ℹ Duplicate skipped — already logged, nothing to do.");
     }
   else if(StringFind(respStr, "\"status\":\"error\"") >= 0 ||
           StringFind(respStr, "\"status\": \"error\"") >= 0)
     {
      Print("TradeJournalEA: ❌ Apps Script returned an error — trade was NOT saved to sheet. ",
            "Full response: ", respStr);
     }
   else if(code == 200)
     {
      // Unexpected response shape — log it clearly
      Print("TradeJournalEA: ⚠ Unexpected response (trade may not have saved): ", respStr);
     }
  }

// ── FORMATTERS ────────────────────────────────────────────────────
string Dbl(double val,int digits)
  {
   if(!MathIsValidNumber(val)||val==EMPTY_VALUE) return "0";
   bool neg=(val<0.0); if(neg) val=-val;
   double scale=MathPow(10.0,digits);
   long whole=(long)val;
   long frac=(long)MathRound((val-(double)whole)*scale);
   if(frac>=(long)scale){whole++;frac=0;}
   string fracStr=""; long tmp=frac;
   for(int i=0;i<digits;i++){fracStr=IntegerToString(tmp%10)+fracStr;tmp/=10;}
   string r=IntegerToString(whole)+"."+fracStr;
   return neg?"-"+r:r;
  }

string Q(string s)    { return "\""+s+"\""; }

string Fmt(datetime dt)
  {
   MqlDateTime t; TimeToStruct(dt,t);
   return StringFormat("%04d-%02d-%02d %02d:%02d",t.year,t.mon,t.day,t.hour,t.min);
  }

string DoW(datetime dt)
  {
   MqlDateTime t; TimeToStruct(dt,t);
   string d[]={"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"};
   return d[t.day_of_week];
  }

string Sess(datetime dt)
  {
   MqlDateTime t; TimeToStruct(dt,t);
   int h=t.hour;
   if(h>=22||h<8)  return "Asia";
   if(h>=8&&h<12)  return "London";
   if(h>=12&&h<17) return "New York";
   if(h>=17&&h<20) return "London Close";
   return "After Hours";
  }

string Dur(datetime open,datetime close)
  {
   if(open==0||close==0) return "";
   int s=(int)MathAbs((double)(close-open));
   return StringFormat("%02dh %02dm",s/3600,(s%3600)/60);
  }

string RR(double entry,double sl,double closeP)
  {
   if(sl==0.0||entry==0.0) return "";
   double risk=MathAbs(entry-sl);
   if(risk==0.0) return "";
   return Dbl(MathAbs(closeP-entry)/risk,2);
  }
//+------------------------------------------------------------------+

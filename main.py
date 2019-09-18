import xlrd, requests, json, time
import warnings
from requests.packages.urllib3.exceptions import InsecureRequestWarning

book = xlrd.open_workbook("dividends.xlsx")
print("Worksheet name(s): {0}".format(book.sheet_names()))
ie_token = '****************************'

class AllStock():
    def __init__(self, row, x):
        self.ticker = row[0].value
        self.company = row[1].value
        self.latestPrice = row[2].value
        self.Dividend_Yield	= row[3].value
        self.Market_Cap = row[4].value
        self.PE_Ratio = row[5].value
        self.Payout_Ratio = row[6].value
        self.location = x

    def funt_init_analyze(self):
        if type(self.Payout_Ratio) is float:
            if self.Payout_Ratio > 55:
                if int(self.PE_Ratio) in range(5, 40):
                    if self.Dividend_Yield > 3.75:
                        if self.Market_Cap > 10:
                            #if int(self.latestPrice) in range(4,12):
                            #print('Possible: ', self.ticker, self.latestPrice)
                            return True

    def funt_stock_basic(self):
        warnings.simplefilter('ignore', InsecureRequestWarning)
        p = requests.get('https://cloud.iexapis.com/stable/stock/' + self.ticker + '/quote?token=' + ie_token, verify=False)
        return json.loads(p.content.decode('utf-8'))

    def funt_financials(self):
        #paid account only
        warnings.simplefilter('ignore', InsecureRequestWarning)
        f = requests.get('https://cloud.iexapis.com/stable/stock/' + self.ticker + '/financials?token=' + ie_token, verify=False)
        financial_lookup = json.loads(f.content.decode('utf-8'))
        return financial_lookup.get('financials')[0]

    def funt_earnings_report(self):
        #paid account only
        warnings.simplefilter('ignore', InsecureRequestWarning)
        earnings = requests.get('https://cloud.iexapis.com/stable/stock/' + self.ticker + '/balance-sheet?token=' + ie_token, verify=False)
        return json.loads(earnings.content.decode('utf-8'))

    def funt_dividends(self):
        warnings.simplefilter('ignore', InsecureRequestWarning)
        d = requests.get('https://cloud.iexapis.com/stable/stock/' + self.ticker + '/dividends/5y?token=' + ie_token, verify=False)
        return json.loads(d.content.decode('utf-8'))

def set_the_att(returned_dict):
    for a in returned_dict:
        try:
            setattr(allstocks, a, returned_dict[a])
        except Exception as e:
            print('line 61', e, returned_dict)

def advanced_research(allstocks):
    if allstocks.latestPrice > (allstocks.week52High * .9):
        pass
    elif allstocks.latestPrice < (allstocks.week52Low * 1.1):
        pass
    elif int(round(allstocks.ytdChange*100)) not in range(0, 30):
        pass
    else:
        return True

sh = [book.sheet_by_index(x) for x in range(0, len(book.sheet_names()))]
block_dups = []
for a in list(set(sh)):
    try:
        for rx in range(a.nrows)[1:]:
            if a.row(rx)[0] not in block_dups:
                allstocks = AllStock(a.row(rx), a)
                block_dups.append(a.row(rx)[0])
                if allstocks.funt_init_analyze():
                    set_the_att(allstocks.funt_stock_basic())
                    if hasattr(allstocks, 'week52Low') and type(allstocks.week52Low) is int or type(allstocks.week52Low) is float:
                        if advanced_research(allstocks):
                            """not available to free acounts
                            #set_the_att(allstocks.funt_financials())
                            if allstocks.netIncome / allstocks.totalAssets > .05:
                                # Not broke
                                if allstocks.operatingIncome - allstocks.operatingExpense < 0:
                                    pass
                                elif allstocks.currentAssets * 1.5 - allstocks.currentDebt < 0:
                                    pass
                                    # retain 10% profit
                                elif allstocks.operatingRevenue / allstocks.operatingIncome < 0.1:
                                    pass
                                else:
                            set_the_att(allstocks.funt_earnings_report())"""
                            set_the_att(allstocks.funt_dividends()[0])
                            print('Success ', allstocks.ticker, allstocks.latestPrice)
                            #get all attributes and values
                            # xx = vars(allstocks).keys()
                            #for x in xx:
                            #    print(x, allstocks.__getattribute__(x))
                            """
                        
                            To do:
                            Analyst opinion
                            Replace financials as financials as reported
                            estimates
                            key stats
                            Sector performance
                        
                            """
    except Exception as e:
        print('line 115 ', a.row(rx)[0], e)
    time.sleep(5)



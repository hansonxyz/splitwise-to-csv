from flask import Flask, render_template, redirect, session, url_for, request
import config as Config
from oauth2client import client
from googlesheets import GoogleSheet
import datetime
import calendar
from splitwise import Splitwise
from apscheduler.schedulers.background import BackgroundScheduler

app = Flask(__name__)
app.secret_key = "test_secret_key"

backupset = False

@app.route("/")
def home():
    global backupset

    if backupset:
        return render_template("home.html")

    #Check for Google Credentials
    if 'googlecredentials' not in session:
        return redirect(url_for('googleLogin'))

    googlecredentials = client.OAuth2Credentials.from_json(session['googlecredentials'])

    if googlecredentials.access_token_expired:
        return redirect(url_for('googleLogin'))

    googleSheet = GoogleSheet(googlecredentials)

    #Check for splitwise credentials
    if 'splitwiseaccesstoken' not in session:
        return redirect(url_for('splitwiseLogin'))

    sObj = Splitwise(Config.splitwise_consumer_key,Config.splitwise_consumer_secret)
    sObj.setAccessToken(session['splitwiseaccesstoken'])

    scheduler = BackgroundScheduler()
    scheduler.add_job(backupData, 'cron', hour=21,minute=34,second=0,args=[googleSheet,sObj])
    scheduler.start()

    backupset = True

    return render_template("home.html")

@app.route("/google/login")
def googleLogin():


    flow = client.OAuth2WebServerFlow(client_id=Config.google_client_id,
    client_secret=Config.google_client_secret,
    scope='https://www.googleapis.com/auth/spreadsheets',
    redirect_uri=url_for('googleLogin', _external=True))

    if 'code' not in request.args:
        auth_uri = flow.step1_get_authorize_url()
        return redirect(auth_uri)
    else:
        auth_code = request.args.get('code')
        credentials = flow.step2_exchange(auth_code)
        session['googlecredentials'] = credentials.to_json()
        return redirect(url_for('home'))

@app.route("/splitwise/login")
def splitwiseLogin():

    if 'splitwisesecret' not in session or 'oauth_token' not in request.args or 'oauth_verifier' not in request.args:
        sObj = Splitwise(Config.splitwise_consumer_key,Config.splitwise_consumer_secret)
        url, secret = sObj.getAuthorizeURL()
        session['splitwisesecret'] = secret
        return redirect(url)

    else:
        oauth_token    = request.args.get('oauth_token')
        oauth_verifier = request.args.get('oauth_verifier')
        sObj = Splitwise(Config.splitwise_consumer_key,Config.splitwise_consumer_secret)
        access_token = sObj.getAccessToken(oauth_token,session['splitwisesecret'],oauth_verifier)
        session['splitwiseaccesstoken'] = access_token
        return redirect(url_for('home'))

    return render_template("home.html")


def backupData(googleSheet,splitwiseObj):
    #Current time
    now = datetime.datetime.now()

    print "Backing up data at "+str(now)

    ########### Get data from Splitwise ####################
    friends = splitwiseObj.getFriends()

    ########## Put data in Google #########################
    spreadsheetName = "SplitwiseBackup"+str(now.year)
    currMonth  = calendar.month_name[now.month]

    #Check if spreadsheet is there if not create a spreadsheet
    if spreadsheetName in Config.spreadsheets:
        spreadsheet = googleSheet.getSpreadSheet(Config.spreadsheets[spreadsheetName])
    else:
        spreadsheet = googleSheet.createSpreadSheet(spreadsheetName,currMonth)
        Config.spreadsheets[spreadsheetName] = spreadsheet.getId()

    #check if current month sheet is there
    sheets = spreadsheet.getSheets()

    sheetPresent = False

    for sheet in sheets:
        if sheet.getName() == currMonth:
            sheetPresent = True
            break

    #If not create a current month sheet
    if not sheetPresent:
        googleSheet.addSheet(spreadsheet.getId(),currMonth)


    #Data to be updated in sheet
    updateData = {
    }


    #Get current sheet data
    data =  googleSheet.getData(spreadsheet.getId(),currMonth+"!A1:Z1000")

    if data is None:
        data = [["Date"]]
        updateData["A1"] = "Date"

    nameRow = data[0]
    newRow = len(data)+1
    lastFilledCol = len(nameRow)

    updateData["A"+str(newRow)]=str(now)

    for friend in friends:
        name = friend.getFirstName()
        amount = ""

        for balance in friend.getBalances():
            amount += balance.getCurrencyCode()+" "+balance.getAmount()+"\n"

        try:#Name is in the list
            index = nameRow.index(name)
        except ValueError as ve: #Name is not in the list
            index = lastFilledCol
            lastFilledCol += 1
            updateData[getColumnNameFromIndex(index)+"1"] = name

        updateData[getColumnNameFromIndex(index)+str(newRow)] = amount

    googleSheet.batchUpdate(spreadsheet.getId(),updateData)

    print "Data backed up successfully"


def getColumnNameFromIndex(index):

    quo = index/26
    rem = index%26

    pre = ""
    post = ""

    if quo != 0:
        pre = chr(quo+64)

    post = chr(rem+65)

    return pre+post


if __name__ == "__main__":
    app.run(threaded=True,debug=Config.debug)

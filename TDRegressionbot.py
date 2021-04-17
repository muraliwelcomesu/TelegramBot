import subprocess,os,openpyxl
from telegram.ext.updater import Updater
from telegram.update import Update
from telegram.ext.callbackcontext import CallbackContext
from telegram.ext.commandhandler import CommandHandler
from telegram.replykeyboardmarkup import ReplyKeyboardMarkup
from telegram.replykeyboardremove import ReplyKeyboardRemove
from telegram.ext.messagehandler import MessageHandler
from telegram.ext.filters import Filters
import TD_BotUtils as utils
import sys,traceback
import Config
token = Config.token

REQUEST_KWARGS={
    # "USERNAME:PASSWORD@" is optional, if you need authentication:
    'proxy_url': Config.proxy_url
}

updater = Updater(token,request_kwargs=REQUEST_KWARGS, use_context=True)


def start(update: Update, context: CallbackContext):
    """
    method to handle the /start command and create keyboard
    """

    # defining the keyboard layout
    kbd_layout = [['TD Specific_J01','TD Specific_GM1'], ['TD Extract','TD Adhoc'],
                       ['J01_STEP1','J01_STEP2'],['J01_STEP3','J01_UPDXL']]

    # converting layout to markup
    
    kbd = ReplyKeyboardMarkup(kbd_layout)

    # sending the reply so as to activate the keyboard
    update.message.reply_text(text="Select Options", reply_markup=kbd)


    
def remove(update: Update, context: CallbackContext):
    """
    method to handle /remove command to remove the keyboard and return back to text reply
    """

    reply_markup = ReplyKeyboardRemove()

    # sending the reply so as to remove the keyboard
    update.message.reply_text(text="I'm back.", reply_markup=reply_markup)
    pass

def help(update: Update, context: CallbackContext):
    lst_commands = []
    lst_commands.append('1. /start - start Program' )
    lst_commands.append('2. /remove -remove keypad' )
    lst_commands.append('3. /getstatus - Get Pending Batch Service Status for branch J01' )
    lst_commands.append('4. /execquery - Execute a Query' )
    lst_commands.append('5. Selecting TD Specific from keypad would trigger Specific Case Execution - rows only in U status will get executed ' )
    lst_commands.append('6. Selecting TD Adhoc   from keypad would trigger Adhoc Execution (all sheets will get executed)' )
    lst_commands.append('7. Selecting J01_UPDXL option will update E Records to U in Specific Case Excel for J01 Branch')
    lst_commands.append('8. Selecting J01_STEP<no>  option will execute the sheet no specified as STEP<no> in Specific Case Excel for J01 Branch')
    
    lst_str = '\n'.join(lst_commands)
    update.message.reply_text(text=lst_str)
    pass

def execquery(update: Update, context: CallbackContext):
    query = str(update.message.text).split('execquery')[1]
    print(query)
    lst_str = utils.execQryReturnStringLst(query)
    if len(lst_str) > 0:
        strresult =  '\n'.join(lst_str)
    else:
        strresult = 'No Records Found for given query'
    update.message.reply_text(text=strresult)
    pass

def getstatus(update: Update, context: CallbackContext):
    l_branch = 'J01' # str(update.message.text.split('_')[1])
    lst_str = utils.fn_get_pending_batch_Service(l_branch)
    update.message.reply_text(text=lst_str)
    pass

def get_error_details(branch):
    try:
        print('inside get_error_details for {}'.format(branch))
        ErrRepPath = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\ErrorReport'
        lst_files = [x for x in os.listdir(ErrRepPath) if x.startswith('{}_ExecLog_'.format(branch))]
        if len(lst_files) > 0:
            lst_files.sort(key=lambda x: os.path.getmtime(os.path.join(ErrRepPath,x)))
            filename = lst_files[-1]
            fp= open(os.path.join(ErrRepPath,filename),'r')
            lst_lines = fp.readlines()
            fp.close()
            lst_Str = ' '.join(lst_lines)
        else:
            lst_Str = '*'
            
        return lst_Str
    except:
        traceback.print_exc()
        return '*'

            
def TDSpecific(update: Update, context: CallbackContext):
    print('Inside TD Specific - {}'.format(update.message.text))
    branch = update.message.text.split('TD Specific_')[1]
    FileInprogress = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\RegressionInProgress_{}.txt'.format(branch)
    if os.path.exists(FileInprogress):
        update.message.reply_text("TD Regression for Branch {}  Execution in Progress.Please wait for Completion".format(branch))
    else:
        
        filepath = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\Specific_Case\{}'.format(branch)
        Lst_Files = os.listdir(filepath)
        Lst_Files = [x for x in Lst_Files if x.endswith('.xlsx') and not x.startswith('~')]
        if len(Lst_Files) > 0 :
            update.message.reply_text("TD Specific Case Execution Started No.of.Files to Process is {}".format(len(Lst_Files)))
            path = r'C:\Murali\TDRegression\BatchFiles\SpecificCases_{}.bat'.format(branch)
            subprocess.call(path)
            update.message.reply_text("TD Specific Regression Completed")
            try:
                strtext = get_error_details(branch)
                if strtext != '*':
                    update.message.reply_text(strtext)
            except:
                print('Some exception')
                traceback.print_exc()
                pass            
        else:
            update.message.reply_text("No Files Found for Processing.Please place testcase excel in ChakraTeam-Share\\Testing_Share\\TDRegression\\Regression_Exec\\Specific_Case\\{} Folder".format(branch))
    pass

def update_excel(ExcelName,p_sheetname):
    wb = openpyxl.load_workbook(ExcelName)
    lst_sheets = sorted([x for x in wb.sheetnames if x.startswith('STEP')])
    for sheet_name  in lst_sheets:
        sheet = wb[sheet_name] 
        for row in range(2,sheet.max_row + 1):
            if sheet_name == p_sheetname:
                sheet['H' + str(row)].value = 'U'
                sheet['J' + str(row)].value = ''
            else:
                sheet['H' + str(row)].value = 'X'
    wb.save(ExcelName)
    

        
def TDSpecificSheet(update: Update, context: CallbackContext):
    i = 'Y'
    if i == 'Y':
        print('Inside TD TDSpecificSheet - {}'.format(update.message.text))
        sheetname = update.message.text.split('J01_')[1]
        branch = 'J01'
        FileInprogress = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\RegressionInProgress_{}.txt'.format(branch)
        if os.path.exists(FileInprogress):
            update.message.reply_text("TD Regression for Branch {}  Execution in Progress.Please wait for Completion".format(branch))
        else:
            print('Sheetname is {}'.format(sheetname))
            ExcelName  = "C:\\ChakraTeam-Share\\Testing_Share\\TDRegression\\Regression_Exec\\Specific_Case\\J01\\J01_TD_Regression1.xlsx"
            try:
                update_excel(ExcelName,sheetname)
                update.message.reply_text('Excel Preparation Complete for Sheet {}. Execution begins now.. '.format(sheetname))
            except:
                update.message.reply_text('Looks like someone has opened specific case excel.. please close it and retry...')
                return 
                
            filepath = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\Specific_Case\{}'.format(branch)
            Lst_Files = os.listdir(filepath)
            Lst_Files = [x for x in Lst_Files if x.endswith('.xlsx') and not x.startswith('~')]
            if len(Lst_Files) > 0 :
                update.message.reply_text("TD Specific Case Execution Started No.of.Files to Process is {}".format(len(Lst_Files)))
                path = r'C:\Murali\TDRegression\BatchFiles\SpecificCases_{}.bat'.format(branch)
                subprocess.call(path)
                update.message.reply_text("TDSpecificSheet Regression Completed for sheet {}".format(sheetname))
                try:
                    strtext = get_error_details(branch)
                    print(strtext)
                    if strtext != '*':
                        update.message.reply_text(strtext)
                except:
                    print('Some exception')
                    traceback.print_exc()
                    pass
                    
            else:
                update.message.reply_text("No Files Found for Processing.Please place testcase excel in ChakraTeam-Share\\Testing_Share\\TDRegression\\Regression_Exec\\Specific_Case\\{} Folder".format(branch))

def TDSpecificUpdExcel(update: Update, context: CallbackContext):
    print('Inside TD TDSpecificUpdExcel - {}'.format(update.message.text))
    branch = 'J01'
    FileInprogress = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\RegressionInProgress_{}.txt'.format(branch)
    if os.path.exists(FileInprogress):
        update.message.reply_text("TD Regression for Branch {}  Execution in Progress.Please wait for Completion".format(branch))
    else:
        ExcelName  = "C:\\ChakraTeam-Share\\Testing_Share\\TDRegression\\Regression_Exec\\Specific_Case\\J01\\J01_TD_Regression1.xlsx"
        wb = openpyxl.load_workbook(ExcelName)
        lst_sheets = sorted([x for x in wb.sheetnames if x.startswith('STEP')])
        for sheet_name  in lst_sheets:
            sheet = wb[sheet_name] 
            for row in range(2,sheet.max_row + 1):
                if sheet['H' + str(row)].value == 'E':
                    sheet['H' + str(row)].value = 'U'
    wb.save(ExcelName)
    update.message.reply_text('Error Record in E status updated to U in Excel.Specific Case can be triggered now..')
               


def TDAdhoc(update: Update, context: CallbackContext):
    FileInprogress = r'C:\ChakraTeam-Share\Testing_Share\TDRegression\Regression_Exec\RegressionInProgress.txt'
    l_brn = 'J01'
    if os.path.exists(FileInprogress):
        update.message.reply_text("TD Regression Execution in Progress.Please wait for Completion")
    else:
        # sending the reply message with the selected option
        #update.message.reply_text("You just clicked on '%s' " % update.message.text)
        update.message.reply_text("TD Adhoc Regression Started")
        path = r'C:\Murali\TDRegression\BatchFiles\Adhoc_Regression1.bat'
        subprocess.call(path)
        update.message.reply_text("TD Adhoc Regression Completed")
        try:
            strtext = get_error_details(l_brn)
            if strtext != '*':
                update.message.reply_text(strtext)
        except:
            print('Some exception')
            traceback.print_exc()
            pass          
    pass

def TDExtract(update: Update, context: CallbackContext):
    """
    message to handle any "Option [0-9]" Regrex.
    """
    update.message.reply_text("Sorry.. Not Yet Implemented..")
    # sending the reply message with the selected option
    #update.message.reply_text("You just clicked on '%s'" % update.message.text)
    pass

def TDExtras(update: Update, context: CallbackContext):
    """
    message to handle any "Option [0-9]" Regrex.
    """
    update.message.reply_text("For Future Implementations...")
    # sending the reply message with the selected option
    #update.message.reply_text("You just clicked on '%s'" % update.message.text)
    pass

updater.dispatcher.add_handler(CommandHandler("start", start))
updater.dispatcher.add_handler(CommandHandler("remove", remove))
updater.dispatcher.add_handler(CommandHandler("help", help))
updater.dispatcher.add_handler(CommandHandler("getstatus",getstatus))
updater.dispatcher.add_handler(CommandHandler("execquery",execquery))
updater.dispatcher.add_handler(MessageHandler(Filters.regex(r"TD Specific"), TDSpecific))
updater.dispatcher.add_handler(MessageHandler(Filters.regex(r"TD Adhoc"), TDAdhoc))
updater.dispatcher.add_handler(MessageHandler(Filters.regex(r"TD Extract"), TDExtract))
updater.dispatcher.add_handler(MessageHandler(Filters.regex(r"<<<<>>>>"), TDExtras))
updater.dispatcher.add_handler(MessageHandler(Filters.regex(r"J01_STEP"), TDSpecificSheet))
updater.dispatcher.add_handler(MessageHandler(Filters.regex(r"J01_UPDXL"), TDSpecificUpdExcel))

print('TD RegressionBot Started..... Please Check with Murali Before Closing This....')

updater.start_polling()

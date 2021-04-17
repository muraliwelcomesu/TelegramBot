 #! python3
import cx_Oracle,os
import Config
#cx_Oracle.init_oracle_client(lib_dir=r"C:\OracleInstantClient\instantclient_19_9")

def fetch_records(p_query):
    p_conn_str = p_conn_str = Config.conn_str
    connection = cx_Oracle.connect(p_conn_str)
    cursor1 = connection.cursor()
    rs = cursor1.execute(p_query)
    result = []
    for i in rs:
        result.append(i)
    cursor1.close()
    connection.close()
    return result

def execQryReturnStringLst(p_query):
    l_conn_str = p_conn_str = Config.conn_str
    result = fetch_records(p_query)
    lst_str  = []
    #l_found = 0 
    for row in result:
        l_row  = list(row)
        l_row = [str(x) for x in l_row]
        #if int(l_found) == 0:
        #    lst_str.append(p_query)
        #    l_found = 1 
        lst_str.append(' '.join(list(l_row)))
    #retstr = '\n'.join(lst_str)
    #print(retstr)
    return lst_str
        

def fn_get_pending_batch_Service(p_brn):
    query = "select service_Seq_no, to_char(processing_date,'DD-MON-YY') processing_date, stage,function_id, service_code ,status ,NVL(err_code,'-') err_code, nvl(err_param,'-') err_param   from  eitb_pending_batch_service where branch_code = '{}' order by service_Seq_no ".format(p_brn)
    result = fetch_records(query)
    lst_final = []
    lst_cols = ['No    ','Date    ','Stage  ','funcId', 'service' ,'status','err_code','err_param']
 
    l_cols = ' '.join(lst_cols)
    lst_final.append(l_cols)
    for i in result:
        tmp_list = [str(x).ljust(6,' ')  for x in list(i)]
        l_vals = ' '.join(tmp_list)
        lst_final.append(l_vals)
    lst_str = '\n'.join(lst_final)
    return lst_str
        
if __name__ == "__main__":
    fn_get_pending_batch_Service('J01')
    

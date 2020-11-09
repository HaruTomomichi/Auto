# 전체건수
def total_sum(t_sum) :
    total = sheet.cell(row=i, column=5).value
    while (total.find(',') > 0):
        total = total.replace(',', '')
    t_sum += int(total)
    return t_sum

# 오류건수
def error_sum(e_sum) :
    error = sheet.cell(row=i, column=6).value
    while (error.find(',') > 0):
        error = error.replace(',', '')
    e_sum += int(error)
    return e_sum

# output
def write(domain, t_sum, e_sum) :
    output.cell(row=j, column=domain, value=t_sum)
    output.cell(row=j, column=domain+1, value=e_sum)
    output.cell(row=j, column=domain+2, value=r_sum)

def sum(col, j) :
    sum = 0
    for row in range (4,j) :
        sum += output.cell(row=row, column=col).value
    return sum

if __name__=="__main__" :

    while True :

        flag = True
        i = 2
        j = 4

        at_sum = bt_sum = ct_sum = dt_sum = et_sum = ft_sum = gt_sum = 0
        ae_sum = be_sum = ce_sum = de_sum = ee_sum = fe_sum = ge_sum = 0

        table = sheet.cell(row=2, column=1).value

        while flag :
            t_read = sheet.cell(row=i, column=1).value
            if (t_read == table) :
                output.cell(row=j, column=1, value=table)

            else:
                if (t_read == None):
                    flag = False
                j = j+1
                table = t_read
                output.cell(row=j, column=1, value=table)
                at_sum = bt_sum = ct_sum = dt_sum = et_sum = ft_sum = gt_sum = 0
                ae_sum = be_sum = ce_sum = de_sum = ee_sum = fe_sum = ge_sum = 0
                r_sum = 0

            d_read = sheet.cell(row=i, column=3).value

            i += 1

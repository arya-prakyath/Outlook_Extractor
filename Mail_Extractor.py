import win32com.client
import xlsxwriter as xl
import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)


importance_list = ['Low', 'Medium', 'High']
sensitivity_list = ['Normal', 'Personal', 'Private', 'Confidential']


def extract_mail():
    try:
        now = datetime.datetime.now().strftime("%d-%m-%y_%H;%M;%S")
        file_name = f"Outlook_Extract_{now}.xlsx"
        workbook = xl.Workbook(file_name)
        worksheet = workbook.add_worksheet()
    except xl.exceptions.FileCreateError:
        print('Some Error Occurred: Close the excel file if open')
        return

    header_style = workbook.add_format({'bold': True})
    worksheet.write('A1', 'Number', header_style)
    worksheet.write('B1', 'Sender', header_style)
    worksheet.write('C1', 'Receiver', header_style)
    worksheet.write('D1', 'CC', header_style)
    worksheet.write('E1', 'Received Date', header_style)
    worksheet.write('F1', 'Importance', header_style)
    worksheet.write('G1', 'Sensitivity', header_style)
    worksheet.write('H1', 'Categories', header_style)
    worksheet.write('I1', 'Read/ Unread', header_style)
    worksheet.write('J1', 'Subject', header_style)
    worksheet.write('K1', 'Body', header_style)

    from_date = input('\nEnter From Date (DD-MM-YYYY)\n')
    to_date = input('Enter To Date (DD-MM-YYYY)\n')
    date_format = "%d-%m-%Y"

    try:
        bool(datetime.datetime.strptime(from_date, date_format))
    except ValueError:
        print('FROM Date is not valid')
        return

    try:
        bool(datetime.datetime.strptime(to_date, date_format))
    except ValueError:
        print('TO Date is not valid')
        return

    from_date = from_date.split('-')
    from_date = datetime.datetime(int(from_date[2]), int(from_date[1]), int(from_date[0]))
    to_date = to_date.split('-')
    to_date = datetime.datetime(int(to_date[2]), int(to_date[1]), int(to_date[0]))

    if to_date < from_date:
        print('FROM Date cannot be greater than the TO Date')
        return

    row = 1
    print('Creating Excel Sheet...')
    messages = inbox.Items
    messages.sort("ReceivedTime", True)
    message = messages.GetFirst()
    while message:
        try:
            received_datetime = str(message.ReceivedTime).split(' ')
            received_date = received_datetime[0]
            received_time = received_datetime[1].split('.')[0]

            temp_received_date = received_date.split('-')
            received_datetime = datetime.datetime(int(temp_received_date[0]),
                                                  int(temp_received_date[1]),
                                                  int(temp_received_date[2]))

            if received_datetime > to_date:
                continue
            if received_datetime < from_date:
                break

            worksheet.write(row, 0, row)

            try:
                worksheet.write(row, 1, str(message.SenderName))
            except Exception:
                worksheet.write(row, 1, '-')

            try:
                worksheet.write(row, 2, str(message.ReceivedByName))
            except Exception:
                try:
                    recipients = ''
                    for recipient in message.Recipients:
                        recipients += str(recipient) + ', '
                    worksheet.write(row, 2, recipients)
                except Exception:
                    worksheet.write(row, 2, '-')

            try:
                worksheet.write(row, 3, str(message.CC))
            except Exception:
                worksheet.write(row, 3, '')

            worksheet.write(row, 4, f'{received_date} {received_time}')

            worksheet.write(row, 5, importance_list[message.Importance])
            worksheet.write(row, 6, sensitivity_list[message.Sensitivity])
            worksheet.write(row, 7, message.Categories)

            worksheet.write(row, 8, ("Unread" if message.UnRead else "Read"))
            worksheet.write(row, 9, str(message.Subject))
            worksheet.write(row, 10, str(message.Body))

            row += 1

        except Exception as e:
            print(f'Error - {e}')

        finally:
            message = messages.GetNext()

    print(f'\nCreated Excel Sheet - {file_name}')
    workbook.close()


extract_mail()
while True:
    run_again = input("\nClick 'Y' and Enter to Run again (or) Click 'Enter' to exit:\n")
    if run_again == 'y' or run_again == 'Y':
        extract_mail()
    else:
        exit(0)

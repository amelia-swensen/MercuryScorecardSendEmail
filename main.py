import get_data
import send_email
import timeit
import check_run_emails

def main():
    start = timeit.default_timer()

    # check_run_emails.read_file()
    check_run_emails.send_outlook_email("amelia.swensen@tigris-fp.com", r"C:\Users\amelia.swensen\OneDrive - TigrisFP\Documents\Accounting\check run email test.csv")
    # send_email.send_outlook_email_hardcoded()

    stop = timeit.default_timer()
    print("Time:", round(stop - start, 2))

if __name__ == '__main__':
    main()

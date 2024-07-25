from time import sleep
import schedule
from main_functions import run_price_monitoring


# schedule the task
schedule.every(30).minutes.do(run_price_monitoring)

print(f'next run will start at: {schedule.next_run}')
while True:
  schedule.run_pending()
  sleep(2)

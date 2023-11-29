import engine
import analysis


def update_data():
    engine.get_data()
    engine.year_data()
    engine.month_data()


analysis.latest_monthly()

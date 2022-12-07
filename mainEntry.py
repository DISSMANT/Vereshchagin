input_value = input('В asd или C: ')

if input_value == 'Вакансии':
    import statisticsChart
    statisticsChart.InputConnect()
elif input_value == 'Статистика':
    import statisticsReport
    statisticsReport.InputConnect()

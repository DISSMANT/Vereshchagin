input_value = input('Вasdансии или Cтатистика: ')

if input_value == 'Вакансии':
    import statisticsChart
    statisticsChart.InputConnect()
elif input_value == 'Статистика':
    import statisticsReport
    statisticsReport.InputConnect()

input_value = input('В или C: ')

if input_value == 'Вакансии':
    import statisticsChart
    statisticsChart.InputConnect()
elif input_value == 'Статистика':
    import statisticsReport
    statisticsReport.InputConnect()

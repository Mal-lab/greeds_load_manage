import pandas as pd
import os
import win32com.client
import matplotlib.pyplot as plt
import math
#Формирует *.csv файлы для каждого узла нагрузки для обеспечения суточного графика в расчетной модели
def transform_powertimeline_to_csv():
    new_tps = {'Номер узла': [], 'P': [], 'Q': []}
    path = 'Расчет'  # Имя папки с основными данными.
    files = os.listdir(path + '/Графики нагрузки')
    tps = pd.read_excel(path + '/Графики нагрузки/' + files[0])
    for j in range(0, len(tps)):
        for i in files:
            tps = pd.read_excel(path + '/Графики нагрузки/' + i)
            new_tps['Номер узла'].append(i[:-5])
            new_tps['P'].append(tps.iloc[j]['P'])
            new_tps['Q'].append(tps.iloc[j]['Q'])
        df = pd.DataFrame(data=new_tps)
        df.to_excel(path + '/Почасовые/' + str(j) + '.xlsx')
        df.to_csv(path + '/csv/' + str(j) + '.csv', index=False, header=False, sep=';')
        new_tps = {'Номер узла': [], 'P': [], 'Q': []}
        return 'Преобразование графиков нагрузки завершено!'
        #Результат сохраняется в папку '/csv' и содержит 24 файла (для каждого часа)

#Построение графиков нагрузки
def load_images():
    path = 'Расчет'
    files = os.listdir(path + '/Графики нагрузки')
    for i in files:
        tps = pd.read_excel(path + '/Графики нагрузки/' + i)
        tps['S'] = (tps['P'] ** 2 + tps['Q'] ** 2) ** 0.5
        fig = tps['S'].plot()
        plt.grid(True)
        plt.savefig(path + '/Рисунки/' + i[:-4] + 'png')
        plt.close('all')
        return 'Изображения графиков нагрузки построены!'




#Расчет электрического режима для следюущих сценариев:
# 1) Режим без учета перегрузочной способности оборудования
# 2) Режим с моделирование действия выключателей при превышении допустимых значений по току
# 3) Режим с гибким регулирование нагрузки
def power_load(regim): # Имя файла режима
    rastr = win32com.client.Dispatch('Astra.Rastr')
    return 'RastrWin успешно загружен'
    path = 'Расчет'  # Имя папки с основными данными.
    files = os.listdir(path + '/csv')
    div = ';'  # Разделитель для последующей корректной записи и чтения .csv файлов
    # Загрузка шаблона RastrWin; Используется шаблон с более точным
    # отображением малых величин. Расположение по умолчанию: 'Документы/RastrWin3/SHABLON/режим.rg2',
    # после установки RastrWin
    shabl = 'C:/Users/pbmal/Documents/RastrWin3/SHABLON/режим.rg2'
    for i in files:
        Rastr = win32com.client.Dispatch('Astra.Rastr')
        print('RastrWin успешно загружен')
        # Загрузка файла режима в Модуль RastrWin. Аналогия функции "Открыть" при обычном использовании RastrWin.
        rg = Rastr.Load(1, path + '/' + regim, shabl)
        vetv = Rastr.Tables('vetv')
        node = Rastr.Tables('node')
        area = Rastr.Tables('area')
        node.ReadCSV(2, path + '/csv/' + i, 'ny,pn,qn', div,
                     'имя = значение')  # Импорт csv файла с нагрузкой в каждом узле по часам
        rg = Rastr.rgm('')  # Расчет режима(Аналогия кнопки "F5" в RastrWin).Данная операция не сохраняет результат
        area.WriteCSV(1, path + '/Потребление/' + i, 'pg,pn,dp,pop', div)
        vetv.WriteCSV(1, path + '/Токовая загрузка ветвей/' + i, 'sta,ip,iq,i_zag', div)
        df = pd.read_csv(path + '/Токовая загрузка ветвей/' + i, sep=';',
                         names=['Состояние ветви', 'Узел начала', 'Узел конца', 'Токовая загрузка,%'])
        dg = df[lambda x: x['Токовая загрузка,%'] >= '1']
        dg['Токовая загрузка,%'] = pd.to_numeric(dg['Токовая загрузка,%'])
        dgg = dg[lambda x: x['Токовая загрузка,%'] >= 100]
        dg = dg[lambda x: x['Токовая загрузка,%'] >= 120]
        dnodes = dgg['Узел конца'].tolist()
        if dg.empty == False:
            # Моделирование выключателей
            dg['Состояние ветви'] = 1
            dg.to_csv(path + '/Перегруженные ветви/' + str(i), index=False, header=False, sep=';')
            vetv.ReadCSV(2, path + '/Перегруженные ветви/' + str(i), 'sta,ip,iq', ';', 'имя = значение')
            rg = Rastr.rgm('')
            area.WriteCSV(1, path + '/Потребление с ограничением/' + i, 'pg,pn,dp,pop', div)
            # Моделирование контроля мощности
            dg['Состояние ветви'] = 0
            dg.to_csv(path + '/Перегруженные ветви/' + str(i), index=False, header=False, sep=';')
            vetv.ReadCSV(2, path + '/Перегруженные ветви/' + str(i), 'sta,ip,iq', ';', 'имя = значение')

            pd_node = pd.read_csv(path + '/csv/' + i, sep=';', names=['Номер узла', 'P', 'Q'])
            pd_node.to_csv(path + '/Нагрузка с контролем/' + str(i), index=False, header=False, sep=';')
            new_node_work = pd.DataFrame(columns=['Состояние', 'Номер узла', 'P', 'Q'])
            for j in dnodes:
                new_node = pd.read_csv(path + '/Нагрузка с контролем/' + str(i), sep=';',
                                       names=['Номер узла', 'P', 'Q'])
                new_node_work = pd.concat([new_node_work, new_node[new_node['Номер узла'] == j]], ignore_index=True)
            new_node_work['Состояние'] = 0
            new_node_work = new_node_work[['Состояние', 'Номер узла', 'P', 'Q']]
            while dg.empty == False:
                new_node_work['P'] = new_node_work['P'] * 0.98
                new_node_work['Q'] = new_node_work['Q'] * 0.98
                new_node_work.to_csv(path + '/Нагрузка с контролем/' + str(i), index=False, header=False, sep=';')
                node.ReadCSV(2, path + '/Нагрузка с контролем/' + i, 'sta,ny,pn,qn', div, 'имя = значение')
                rg = Rastr.rgm('')
                vetv.WriteCSV(1, path + '/Ветви с контролем/' + i, 'sta,ip,iq,i_zag', div)
                df = pd.read_csv(path + '/Ветви с контролем/' + i, sep=';',
                                 names=['Состояние ветви', 'Узел начала', 'Узел конца', 'Токовая загрузка,%'])
                dg = df[lambda x: x['Токовая загрузка,%'] >= '1']
                dg['Токовая загрузка,%'] = pd.to_numeric(dg['Токовая загрузка,%'])
                dg = dg[
                    lambda x: x['Токовая загрузка,%'] >= 120]  # >= Значения при котором начинает работу реле мощности
            area.WriteCSV(1, path + '/Потребление с контролем мощности/' + i, 'pg,pn,dp,pop', div)
        else:
            area.WriteCSV(1, path + '/Потребление с ограничением/' + i, 'pg,pn,dp,pop', div)
            area.WriteCSV(1, path + '/Потребление с контролем мощности/' + i, 'pg,pn,dp,pop', div)
    return 'Расчет режимов завершен!'
#Построение графиков и краткая информация об эффективности применения устройств гибкого регулирования нагрузки для конкретной сети
def make_image():
    new_tps = {'№': [], 'Pген': [], 'Pнаг': [], 'Потери': [], 'Pпотр': []}
    path = 'Расчет'
    files = os.listdir(path + '/Потребление')

    for i in files:
        tps = pd.read_csv(path + '/Потребление/' + i, sep=';', names=['Pген', 'Pнаг', 'Потери', 'Pпотр'])
        new_tps['№'].append(int(i[:-4]))
        new_tps['Pген'].append(tps.iloc[0]['Pген'])
        new_tps['Pнаг'].append(tps.iloc[0]['Pнаг'])
        new_tps['Потери'].append(tps.iloc[0]['Потери'])
        new_tps['Pпотр'].append(tps.iloc[0]['Pпотр'])
    df = pd.DataFrame(data=new_tps)
    df['№'] = pd.to_numeric(df['№'])
    df = df.sort_values(by=['№'])


    files1 = os.listdir(path + '/Потребление с ограничением')
    new_tps1 = {'№': [], 'Pген': [], 'Pнаг': [], 'Потери': [], 'Pпотр': []}
    for i in files:
        tps1 = pd.read_csv(path + '/Потребление с ограничением/' + i, sep=';',
                           names=['Pген', 'Pнаг', 'Потери', 'Pпотр'])
        new_tps1['№'].append(int(i[:-4]))
        new_tps1['Pген'].append(tps1.iloc[0]['Pген'])
        new_tps1['Pнаг'].append(tps1.iloc[0]['Pнаг'])
        new_tps1['Потери'].append(tps1.iloc[0]['Потери'])
        new_tps1['Pпотр'].append(tps1.iloc[0]['Pпотр'])
    dg = pd.DataFrame(data=new_tps1)
    dg['№'] = pd.to_numeric(df['№'])
    dg = dg.sort_values(by=['№'])

    files2 = os.listdir(path + '/Потребление с контролем мощности')
    new_tps2 = {'№': [], 'Pген': [], 'Pнаг': [], 'Потери': [], 'Pпотр': []}
    for i in files:
        tps1 = pd.read_csv(path + '/Потребление с контролем мощности/' + i, sep=';',
                           names=['Pген', 'Pнаг', 'Потери', 'Pпотр'])
        new_tps2['№'].append(int(i[:-4]))
        new_tps2['Pген'].append(tps1.iloc[0]['Pген'])
        new_tps2['Pнаг'].append(tps1.iloc[0]['Pнаг'])
        new_tps2['Потери'].append(tps1.iloc[0]['Потери'])
        new_tps2['Pпотр'].append(tps1.iloc[0]['Pпотр'])
    dn = pd.DataFrame(data=new_tps2)
    dn['№'] = pd.to_numeric(dn['№'])
    dn = dn.sort_values(by=['№'])

    total = df['Pнаг'].tolist()
    total1 = dg['Pнаг'].tolist()
    total2 = dn['Pнаг'].tolist()

    print('Значения потребления за сутки:')
    print('Без управляющих воздействий: ', round(sum(total), 2))
    print('При отключении перегруженных линий (>120%): ', round(sum(total1), 2))
    print('При гибком регулировании нагрузки: ', round(sum(total2), 2))
    print('Таким образом недоотпуск эдектроэнергии при отключении линий равен: ', round(sum(total) - sum(total1), 2), ' Мвт*ч. ',
          "Процент выдачи мощности от максимальной : ", round((1 - (sum(total) - sum(total1)) / sum(total)) * 100, 2), ' %')
    print('Таким образом недоотпуск эдектроэнергии при гибком регулировании нагрузки: ', round(sum(total) - sum(total2), 2),
          ' Мвт*ч. ', "Процент выдачи мощности от максимальной : ", round((1 - (sum(total) - sum(total2)) / sum(total)) * 100, 2), ' %')

    plt.bar(df['№'], df['Pнаг'],
            color='orange',
            linestyle='solid',
            label='Pнаг')

    plt.plot(dn['№'], dn['Pнаг'],
             color='red',
             linestyle='dashed',
             label='Pнаг(Контроль P)'
             )

    plt.plot(dg['№'], dg['Pнаг'],
             color='green',
             linestyle='dashed',
             label='Pнаг(Выключатели)')

    plt.legend(loc='upper left')
    plt.grid(True)
    plt.savefig(path + '/Итоговый график.png')
    return 'Итоговый график построен'

print(transform_powertimeline_to_csv())
print(load_images())
print(power_load('режим КП.rg2'))
print(make_image())
print('Полный цикл программы успешно выполнен')
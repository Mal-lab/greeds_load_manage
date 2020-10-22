# greeds_load_manage
Моделирование гибкого регулирования нагрузки в электрической сети с использованием ПК RastrWin3 и Python.
# Назначение
Проект выполнен для демонстрации преимущества использования реле приоритета нагрузок для оптимизации передачи электроэнергии.
Позволяет провести первичный анализ использования для заранее смоделированной расчетной модели в программной среде RastrWin3.
# Принцип работы
Для каждого узла расчетной модели задается суточный график нагрузки. Графики составлены таким образом, чтобы в разные часы
суток возникала перегрузка воздушных линий электропередач. Существующая практика предполагает отключение перегружаемых элементов,
что приводит к неодоотпуску электроэнергии. Реле приоритета нагрузок позволяет ограничивать неприоритетных потребителей, тем самым
сохраняя электроснабжение остальных потребителей.
# О ПК RastrWin3
Широкоприменяемая программа для расчета параметров электрических режимов.
Для оценки эффективности применения гибкого регулирования нагрузки необходим расчет режима для каждого часа с реализацией
логики регулирования и отключения перегружаеммых элементов. Стаданртные средства RastrWin3 не предусматривают подобного
применения, поэтому для решения задачи использован функционал Python с применением библиотек: Pandas, Matplotlib.
# Требования для работы программы:
 1) Необходима предварительная установка ПК RastrWin3, так как для реализации расчета режима используется COM-объект;
 2) Python 3.X 32 bit (64-битная версия не поддерживается)

# Результат выполнения программы
![Alt text](https://github.com/Mal-lab/greeds_load_manage/blob/main/%D0%98%D1%82%D0%BE%D0%B3%D0%BE%D0%B2%D1%8B%D0%B9%20%D0%B3%D1%80%D0%B0%D1%84%D0%B8%D0%BA.png)

Значения потребления за сутки:

Без управляющих воздействий:  18.18 МВт*ч

При отключении перегруженных линий (>120%):  16.36 МВт*ч

При гибком регулировании нагрузки:  17.91 МВт*ч

Таким образом, недоотпуск электроэнергии при отключении линий равен:  1.82  МВт*ч.  Процент выдачи мощности от максимальной:  89.99  %

Таким образом, недоотпуск электроэнергии при гибком регулировании нагрузки:  0.27  МВт*ч.  Процент выдачи мощности от максимальной:  98.53  %





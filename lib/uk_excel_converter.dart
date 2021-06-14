import 'dart:io';

import 'package:excel/excel.dart';
import 'package:in_date_utils/in_date_utils.dart';
import 'package:list_ext/list_ext.dart';
import 'package:path/path.dart' as p;

const _titlePrefix = 'Начисления на лицевые счета';
const _lsPrefix = 'л/с №';
const _months = [
  'Январь',
  'Февраль',
  'Март',
  'Апрель',
  'Май',
  'Июнь',
  'Июль',
  'Август',
  'Сентябрь',
  'Октябрь',
  'Ноябрь',
  'Декабрь'
];
const _typeValue = 0;
const _typePeni = 2;
const _defaultService = 3201;

final _monthRegEx =
    RegExp('за (.*) ([0-9]{4}) г\.', caseSensitive: false, unicode: true);
final _serviceMap = <String, int>{
  'Взнос на капитальный ремонт': 201,
  'Задолженность по взносам ФКР': 299,
  'Задолженность по МУП РИЦ на 1.02.15г.': 199,
  'Закупка материала на ремонт кровли (Победы 2)': 3273,
  'Закупка материала на ремонт крыши (АЭР 28)': 3273,
  'Замена лежака ГВС (ул. Дзержинского 1)': 3272,
  'Замена лежака отопления': 3234,
  'Замена лежака ХВС (АЭР № 28)': 3271,
  'Замена лежака ХВС (АЭР № 6)': 3271,
  'Коммунальный ресурс: ГВС': 189,
  'КР на СОИ: ГВС - Тепловая энергия': 179,
  'КР на СОИ: ГВС - Холодная вода': 176,
  'КР на СОИ: ХВС': 175,
  'КР на СОИ: ХВС (кв. 37,40,42,43,46)': 175,
  'КР на СОИ: ХВС (полив)': 584,
  'КР на СОИ: Электроэнергия': 180,
  'Обслуживание спецсчета (капремонт)': 3249,
  'Перерасчет': 199,
  'Поверка ОДПУ тепла  (2020)': 3203,
  'Поверка ОДПУ тепла - (2018)': 3203,
  'Поверка ОДПУ тепла - (2019)': 3203,
  'Поверка ОДПУ тепла - (2020)': 3203,
  'Поверка ОДПУ тепла (2018)': 3203,
  'Поверка счетчика учета тепла (АЭР 3)': 3203,
  'Поверка счетчика учета тепла (ул. Аэродромная 14)': 3203,
  'Поверка счетчика учета тепла (ул. Советская 121)': 3203,
  'Поверка счетчика учета тепла, Победы 60': 3203,
  'Приобретение и установка счетчика ХВС (Тер. 22а)': 3208,
  'Приобретение ОДПУ ГВС': 3208,
  'Приобретение счетчика ХВС (ул. Аэродромная 4)': 3208,
  'Приобретение/установка Вычислителя ОДПУ': 3209,
  'Ремонт входа в подвал': 3275,
  'Ремонт крыши МКД (Чкалова 1)': 3269,
  'Содержание коменданта и консьержей': 112,
  'Содержание коменданта и швейцар-уборщиц': 111,
  'Содержание уборщиц и консьержей': 113,
  'Текущий ремонт и содержание дома': 102,
  'Текущий ремонт и содержание дома (общежитие)': 114,
  'Текущий ремонт и содержание дома, без ТКО (178)': 115,
  'Текущий ремонт и содержание дома, в т.ч. ТКО': 110,
  'Установка датчиков движения': 3270,
  'Установка пластиковых окон': 3274,
};

Future<void> convert(String filePath, String outputDirPath) async {
  final bytes = File(filePath).readAsBytesSync();
  final original = Excel.decodeBytes(bytes);

  final originaName = p.basename(filePath);
  final source = original.tables.values.first;
  final out = _OutExporter(originaName);

  DateTime? month;
  var service = _defaultService;

  for (final row in source.rows) {
    final first = row.first?.value as String?;
    if (first == null) continue;

    if (month == null) {
      if (first.startsWith(_titlePrefix)) {
        // Начисления на лицевые счета  за Октябрь 2020 г.
        final matches = _monthRegEx.allMatches(first);
        final match = matches.first;
        if (match.groupCount == 2) {
          final monthVal = match.group(1) as String;
          final yearVal = match.group(2) as String;
          if (!_months.contains(monthVal)) {
            throw Exception('Не найден месяц: $monthVal');
          }

          month = DateTime(int.parse(yearVal), _months.indexOf(monthVal) + 1);
        } else {
          continue;
        }
      } else {
        continue;
      }
    }

    if (first.startsWith(_lsPrefix)) {
      final lsNum = int.parse(first.replaceFirst(_lsPrefix, ''));
      final date = _date(DateUtils.firstDayOfNextMonth(month));
      final dDate = _date(month);

      final res = [lsNum, date, dDate, service];

      final value = row[7]?.value as num?;
      final peni = row[8]?.value as num?;

      if (value != null) {
        final r = res.copyWith(value)..add(_typeValue);
        out.appendRow(r);
      }

      if (peni != null) {
        final r = res.copyWith(peni)..add(_typePeni);
        out.appendRow(r);
      }
    } else {
      service = _serviceMap[first] ?? _defaultService;
    }
  }

  final res = await Future.wait([
    out.save(outputDirPath),
  ]);

  print('Данные записаны в файлы:\n${res.map((f) => f.path).join('\n')}');
}

void loadServiceMapFromExcel(String filePath) {
  final bytes = File(filePath).readAsBytesSync();
  final original = Excel.decodeBytes(bytes);

  final source = original.tables.values.first;

  _serviceMap.clear();

  for (final row in source.rows) {
    final key = row.first?.value as String?;
    final value = row[1]?.value;

    if (key != null && value != null && value is int) {
      print("'$key': $value,");
      _serviceMap[key] = value;
    }
  }
}

String _date(DateTime value) =>
    '${_num(value.day)}.${_num(value.month)}.${_num(value.year, 4)}';
String _num(int value, [int digits = 2]) =>
    value.toString().padLeft(digits, '0');

// const prefix = 'Исх';
// final targetPath = p.join(p.dirname(input), '$prefix${p.basename(input)}');

class _OutExporter extends _Exporter {
  _OutExporter(String originalName) : super('Исх', originalName);

  @override
  void appendHeaders() {
    appendRow(['LS', 'MONTH', 'D_MONTH', 'CD_SRV', 'S_SALDO', 'TIP']);
  }
}

abstract class _Exporter {
  final String prefix;
  final String originalName;

  late final Excel excel;

  _Exporter(this.prefix, this.originalName) {
    excel = Excel.createExcel();
    appendHeaders();
  }

  void appendHeaders();

  Future<File> save(String outputDir) async {
    final targetPath = p.join(outputDir, '$prefix$originalName');
    final file = File(targetPath);
    await file.writeAsBytes(excel.save()!);
    return file;
  }

  void appendRow(List<Object> row) => excel.sheets.values.first.appendRow(row);
}

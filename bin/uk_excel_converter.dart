import 'package:uk_excel_converter/uk_excel_converter.dart';
import 'package:path/path.dart' as p;

void main(List<String> arguments) async {
  // loadServiceMapFromExcel(
  // '/Volumes/MacbookExt/projects/dart/uk_excel_converter/Соответствие услуг.xlsx');
  final input =
      '/Volumes/MacbookExt/projects/dart/uk_excel_converter/Октябрь 2020.xlsx';
  final lsMapFilePath =
      '/Volumes/MacbookExt/projects/dart/uk_excel_converter/Соответствие ЛС и СЧЕТА.xlsx';
  final providersMapFilePath =
      '/Volumes/MacbookExt/projects/dart/uk_excel_converter/Номер поставщика для КАП_РЕМОНТА ВСЕ.xlsx';
  await convert(
    input,
    p.dirname(input),
    lsMapFilePath: lsMapFilePath,
    providersMapFilePath: providersMapFilePath,
  );
}

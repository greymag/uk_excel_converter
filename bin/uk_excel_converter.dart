import 'package:uk_excel_converter/uk_excel_converter.dart';
import 'package:path/path.dart' as p;

void main(List<String> arguments) async {
  // loadServiceMapFromExcel(
  // '/Volumes/MacbookExt/projects/dart/uk_excel_converter/Соответствие услуг.xlsx');
  final input =
      '/Volumes/MacbookExt/projects/dart/uk_excel_converter/Октябрь 2020.xlsx';
  const prefix = 'Исх';
  final targetPath = p.join(p.dirname(input), '$prefix${p.basename(input)}');
  final file = await convert(input, targetPath);
  print(file.path);
}

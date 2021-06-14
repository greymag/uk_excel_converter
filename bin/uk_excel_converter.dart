import 'dart:io';

import 'package:uk_excel_converter/uk_excel_converter.dart';
import 'package:path/path.dart' as p;

void main(List<String> arguments) async {
  if (arguments.length != 4) {
    _usage();
    return;
  }

  final sourceDirPath = arguments[0];
  final targetDirPath = arguments[1];
  final lsMapFilePath = arguments[2];
  final providersMapFilePath = arguments[3];

  final dir = Directory(sourceDirPath);
  await for (final file in dir.list()) {
    if (file is File && p.extension(file.path) == '.xlsx') {
      print('Exporting ${p.basename(file.path)}');
      final path = file.path;
      await convert(
        path,
        targetDirPath,
        lsMapFilePath: lsMapFilePath,
        providersMapFilePath: providersMapFilePath,
      );
    }
  }

  print('All done');
}

void _usage() => print('Запустите команду со следующими аргументами:\n'
    '- Путь до директории с файлами для экспорта\n'
    '- Путь до директории куда будут сохранены экспортированные файлы\n'
    '- Путь до файла соответствия ЛС и счета\n'
    '- Путь до файла поставщиков');

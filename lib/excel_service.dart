import 'dart:io';
import 'package:open_file/open_file.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:syncfusion_officechart/officechart.dart';

class ExcelService {
  final List<List<dynamic>>? _excelData;
  final String? _filePath;
  final List<Map<String, dynamic>> _events = [];
  String? _savedFilePath;

  ExcelService(this._filePath, this._excelData);

  Future<void> createExcelTables() async {
    if (_excelData == null || _excelData.isEmpty) {
      print('No data available to create events.');
      return;
    }

    _events.clear();

    for (int i = 1; i < _excelData.length; i++) {
      final entry = _excelData[i]; // Use safe access to avoid null
      var custId = entry.isNotEmpty ? entry[0] : null;
      var arrivalTime = entry.length > 2 ? entry[2] : null;
      var departureTime = entry.length > 7 ? entry[7] : null;

      if (custId != null && arrivalTime != null && departureTime != null) {
        try {
          if (_isNumeric(custId.toString()) &&
              _isNumeric(arrivalTime.toString()) &&
              _isNumeric(departureTime.toString())) {
            int custNum = int.parse(custId.toString());
            int arrival = int.parse(arrivalTime.toString());
            int departure = int.parse(departureTime.toString());

            _events.add({
              'Event_Type': 'Arrival',
              'Cust_Num': custNum,
              'Clock_Time': arrival,
            });
            _events.add({
              'Event_Type': 'Departure',
              'Cust_Num': custNum,
              'Clock_Time': departure,
            });
          } else {
            print('Non-numeric data found in entry: $entry');
          }
        } catch (e) {
          print('Error parsing values: $e');
        }
      } else {
        print('Missing data in entry: $entry');
      }
    }
    await saveEventsToExcel();
  }

  bool _isNumeric(String str) {
    return int.tryParse(str) != null;
  }

  Future<void> saveEventsToExcel() async {
    // Request storage permission
    await _requestStoragePermission();

    final workbook = xlsio.Workbook();
    final sheet = workbook.worksheets[0];
    print('@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ first table');

    // Define headers for the simulationData table.
    final headers = [
      'Cust_id',
      'Interval',
      'Arr.Clock',
      'code',
      'Service',
      'Start',
      'Duration',
      'End_Clock',
      'State Ser',
      'Cust Wait'
    ];

    // Insert headers into the first row.
    for (int i = 0; i < headers.length; i++) {
      sheet.getRangeByIndex(1, i + 1).setText(headers[i]);
    }

    // Insert simulationData into the sheet starting from row 2.
    for (int i = 0; i < _excelData!.length; i++) {
      final row = _excelData[i]; // Use safe access
      for (int j = 0; j < row.length; j++) {
        sheet.getRangeByIndex(i + 2, j + 1).setText(row[j].toString());
      }
    }

    print('============================================ first table');

    // Second Table
    sheet.getRangeByName('A16').setText('Event_Type');
    sheet.getRangeByName('B16').setText('Cust_Num');
    sheet.getRangeByName('C16').setText('Clock_Time');

    for (int i = 0; i < _events.length; i++) {
      final event = _events[i];
      sheet.getRangeByName('A${i + 17}').setText(event['Event_Type']);
      sheet
          .getRangeByName('B${i + 17}')
          .setNumber(event['Cust_Num'].toDouble());
      sheet
          .getRangeByName('C${i + 17}')
          .setNumber(event['Clock_Time'].toDouble());
    }

    // Create chart
    final ChartCollection charts = ChartCollection(sheet);
    final Chart chart = charts.add();
    chart.chartType = ExcelChartType.column;
    chart.dataRange = sheet.getRangeByName('A16:C${16 + _events.length}');

    // Set the position of the chart
    chart.topRow = 16;
    chart.bottomRow = 27; // Adjust this as needed
    chart.leftColumn = 5;
    chart.rightColumn = 16;

    // Set charts to worksheet.
    sheet.charts = charts;

    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    print('============================================ second table');

    // Use the path_provider to get a valid writable path
    final directory = await getApplicationDocumentsDirectory();
    final savedFile = File('${directory.path}/_events.xlsx');

    try {
      await savedFile.writeAsBytes(bytes);
      _savedFilePath = savedFile.path;

      print(
          '############################################# Events saved to Excel at: $_savedFilePath');
    } catch (e) {
      print('Error saving Excel file: $e');
    }
  }

  Future<void> _requestStoragePermission() async {
    var status = await Permission.storage.status;
    if (!status.isGranted) {
      await Permission.storage.request();
    }
  }

  Future<void> openSavedFile() async {
    if (_savedFilePath != null) {
      final result = await OpenFile.open(_savedFilePath);
      if (result.message != 'Done') {
        print('Failed to open file: ${result.message}');
      }
    } else {
      print('No file path available to open.');
    }
  }

  List<Map<String, dynamic>> get events => _events;
  String? get filePath => _filePath;
  String? get savedFilePath => _savedFilePath;
}

import 'dart:io';
import 'package:flutter/foundation.dart';
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
      if (kDebugMode) {
        print('No data available to create events.');
      }
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
            if (kDebugMode) {
              print('Non-numeric data found in entry: $entry');
            }
          }
        } catch (e) {
          if (kDebugMode) {
            print('Error parsing values: $e');
          }
        }
      } else {
        if (kDebugMode) {
          print('Missing data in entry: $entry');
        }
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

    // Style for header cells: yellow background, center alignment, and border
    final headerStyle = workbook.styles.add('HeaderStyle');
    headerStyle.backColor = '#FFFF00'; // Yellow color
    headerStyle.hAlign = xlsio.HAlignType.center;
    headerStyle.bold = true;
    headerStyle.borders.all.lineStyle =
        xlsio.LineStyle.thin; // Border for header cells

    // Style for content cells: center alignment and border
    final contentStyle = workbook.styles.add('ContentStyle');
    contentStyle.hAlign = xlsio.HAlignType.center;
    contentStyle.borders.all.lineStyle =
        xlsio.LineStyle.thin; // Border for content cells

    // Insert headers into the first row with header style.
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(1, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle = headerStyle; // Apply header style
    }

    // Insert simulationData into the sheet starting from row 2 with content style.
    for (int i = 0; i < _excelData!.length; i++) {
      final row = _excelData[i];
      for (int j = 0; j < row.length; j++) {
        final cell = sheet.getRangeByIndex(i + 2, j + 1);
        cell.setText(row[j].toString());
        cell.cellStyle = contentStyle; // Apply content style with borders
      }
    }

    // Second Table with headers
    sheet.getRangeByName('A9').setText('Event_Type');
    sheet.getRangeByName('B9').setText('Cust_Num');
    sheet.getRangeByName('C9').setText('Clock_Time');
    sheet.getRangeByName('A9:C9').cellStyle =
        headerStyle; // Apply header style to second table headers

    for (int i = 0; i < _events.length; i++) {
      final event = _events[i];
      final cellA = sheet.getRangeByName('A${i + 10}');
      cellA.setText(event['Event_Type']);
      cellA.cellStyle = contentStyle;

      final cellB = sheet.getRangeByName('B${i + 10}');
      cellB.setNumber(event['Cust_Num'].toDouble());
      cellB.cellStyle = contentStyle;

      final cellC = sheet.getRangeByName('C${i + 10}');
      cellC.setNumber(event['Clock_Time'].toDouble());
      cellC.cellStyle = contentStyle;
    }

    // Create chart
    final ChartCollection charts = ChartCollection(sheet);
    final Chart chart = charts.add();
    chart.chartType = ExcelChartType.column;
    chart.dataRange = sheet.getRangeByName('A9:C${9 + _events.length}');

    // Set the position of the chart
    chart.topRow = 9;
    chart.bottomRow = 18; // Adjust this as needed
    chart.leftColumn = 5;
    chart.rightColumn = 16;

    // Add charts to worksheet
    sheet.charts = charts;

    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    // Save file in application documents directory
    final directory = await getApplicationDocumentsDirectory();
    final savedFile = File('${directory.path}/_events.xlsx');

    try {
      await savedFile.writeAsBytes(bytes);
      _savedFilePath = savedFile.path;

      if (kDebugMode) {
        print('Events saved to Excel at: $_savedFilePath');
      }
    } catch (e) {
      if (kDebugMode) {
        print('Error saving Excel file: $e');
      }
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
        if (kDebugMode) {
          print('Failed to open file: ${result.message}');
        }
      }
    } else {
      if (kDebugMode) {
        print('No file path available to open.');
      }
    }
  }

  List<Map<String, dynamic>> get events => _events;
  String? get filePath => _filePath;
  String? get savedFilePath => _savedFilePath;
}

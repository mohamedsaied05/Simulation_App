import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;

class ExcelProbService {
  final List<List<dynamic>> _analysisData;
  final List<List<dynamic>> _newSimulationData;
  static String? _savedFilePath;

  ExcelProbService(this._analysisData, this._newSimulationData);

  Future<void> saveExcelProbTables() async {
    // Request storage permission
    await _requestStoragePermission();

    // Create a new Excel workbook
    final xlsio.Workbook workbook = xlsio.Workbook();
    final xlsio.Worksheet sheet = workbook.worksheets[0];

    // Set headers for Analysis Data
    final headers = [
      'Cust_id',
      'ServType',
      'Avg.Dur',
      'Prob',
      'Cum.Prob',
      'From',
      'To'
    ];

    // Define cell style for headers
    final headerStyle = workbook.styles.add('HeaderStyle');
    headerStyle.backColor = '#FFFF00'; // Yellow color
    headerStyle.hAlign = xlsio.HAlignType.center; // Center alignment
    headerStyle.bold = true; // Make headers bold

    // Set headers and apply style
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(1, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle = headerStyle; // Apply header style
      cell.cellStyle.borders.all.lineStyle =
          xlsio.LineStyle.thin; // Add borders to header cells
    }

    // Insert analysis data rows if available
    for (int i = 0; i < _analysisData.length; i++) {
      for (int j = 0; j < _analysisData[i].length; j++) {
        if (j < headers.length) {
          final cell = sheet.getRangeByIndex(i + 2, j + 1);
          cell.setText(_analysisData[i][j]?.toString() ?? '');
          cell.cellStyle.hAlign =
              xlsio.HAlignType.center; // Center alignment for data cells
          cell.cellStyle.borders.all.lineStyle =
              xlsio.LineStyle.thin; // Add borders to data cells
        }
      }
    }

    // Leave a separator row and then add headers for Simulation Data
    int simulationDataStartIndex = _analysisData.length + 3;
    final simulationHeaders = [
      'Cust_id',
      'Interval',
      'Arr.Clock',
      'Code',
      'Service',
      'Start',
      'Duration',
      'End.Clock',
      'State',
      'Cust.Wait'
    ];

    // Define cell style for simulation headers
    final simulationHeaderStyle = workbook.styles.add('SimulationHeaderStyle');
    simulationHeaderStyle.backColor = '#FFFF00'; // Yellow color
    simulationHeaderStyle.hAlign = xlsio.HAlignType.center; // Center alignment
    simulationHeaderStyle.bold = true; // Make headers bold

    // Set headers for simulation data and apply style
    for (int i = 0; i < simulationHeaders.length; i++) {
      final cell = sheet.getRangeByIndex(simulationDataStartIndex, i + 1);
      cell.setText(simulationHeaders[i]);
      cell.cellStyle = simulationHeaderStyle; // Apply header style
      cell.cellStyle.borders.all.lineStyle =
          xlsio.LineStyle.thin; // Add borders to header cells
    }

    // Insert simulation data rows if available
    for (int i = 0; i < _newSimulationData.length; i++) {
      for (int j = 0; j < _newSimulationData[i].length; j++) {
        if (j < simulationHeaders.length) {
          final cell =
              sheet.getRangeByIndex(simulationDataStartIndex + 1 + i, j + 1);
          cell.setText(_newSimulationData[i][j]?.toString() ?? '');
          cell.cellStyle.hAlign =
              xlsio.HAlignType.center; // Center alignment for data cells
          cell.cellStyle.borders.all.lineStyle =
              xlsio.LineStyle.thin; // Add borders to data cells
        }
      }
    }

    // Save the workbook and dispose of it
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

  String? get savedFilePath => _savedFilePath;
}

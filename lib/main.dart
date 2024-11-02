import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'dart:math';
import 'package:simulation_app/excel_service.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return const MaterialApp(
      home: ExcelSimulationScreen(),
    );
  }
}

class ExcelSimulationScreen extends StatefulWidget {
  const ExcelSimulationScreen({super.key});

  @override
  _ExcelSimulationScreenState createState() => _ExcelSimulationScreenState();
}

class _ExcelSimulationScreenState extends State<ExcelSimulationScreen> {
  List<List<String>> excelData = [];
  List<List<String>> simulationData = [];
  bool isExcelLoaded = false;
  bool isExcelGenerated = false;
  String? _filePath = '';
  late ExcelService service;

  @override
  void initState() {
    super.initState();
    service = ExcelService(_filePath, simulationData);
  }

  Future<void> pickAndReadExcelFile() async {
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
      );

      if (result != null) {
        _filePath = result.files.single.path; // Save the file path
        var bytes = File(_filePath!).readAsBytesSync();
        var excel = Excel.decodeBytes(bytes);

        if (excel.tables.isNotEmpty) {
          String firstTableName = excel.tables.keys.first;
          var rows = excel.tables[firstTableName]?.rows;

          if (rows != null && rows.isNotEmpty) {
            excelData = rows
                .map((row) =>
                    row.map((cell) => cell?.value?.toString() ?? '').toList())
                .toList();
            isExcelLoaded = true;
          } else {
            excelData = [
              ['Error:', 'The Excel sheet is empty or invalid.']
            ];
            isExcelLoaded = false;
          }
        }
      }
    } catch (e) {
      excelData = [
        ['Error:', e.toString()]
      ];
      isExcelLoaded = false;
    }
    setState(() {});
  }

  void runSimulation() {
    if (excelData.isEmpty || excelData.length <= 1) return;

    simulationData.clear();
    int currentTime = 0;
    Random random = Random();

    for (int i = 1; i <= 5; i++) {
      int interval = random.nextInt(3) + 1;
      currentTime += interval;
      int codeIndex = random.nextInt(excelData.length - 1) + 1;
      var serviceRow = excelData[codeIndex];

      int duration = int.tryParse(serviceRow[2]) ?? 0;
      int start = max(currentTime,
          simulationData.isNotEmpty ? int.parse(simulationData.last[7]) : 0);
      int endClock = start + duration;
      int custWait = start - currentTime;

      simulationData.add([
        i.toString(),
        interval.toString(),
        currentTime.toString(),
        serviceRow[0],
        serviceRow[1],
        start.toString(),
        duration.toString(),
        endClock.toString(),
        i == 1 ? 'Wait' : 'Busy',
        custWait.toString(),
      ]);
    }
    setState(() {});
  }

  Widget buildTable(List<List<String>> data, List<String> headers) {
    int columnCount = headers.length;

    return SingleChildScrollView(
      scrollDirection: Axis.horizontal,
      child: SingleChildScrollView(
        scrollDirection: Axis.vertical,
        child: Table(
          border: TableBorder.all(color: Colors.black),
          defaultColumnWidth: const FixedColumnWidth(100.0),
          children: [
            TableRow(
              children: headers
                  .map((header) => _buildCell(header, isHeader: true))
                  .toList(),
            ),
            ...data.map((row) {
              while (row.length < columnCount) {
                row.add(''); // Fill missing cells with empty strings
              }
              return TableRow(
                children: row.map((cell) => _buildCell(cell)).toList(),
              );
            }),
          ],
        ),
      ),
    );
  }

  Widget _buildCell(String content, {bool isHeader = false}) {
    return Container(
      padding: const EdgeInsets.all(8.0),
      color: isHeader ? Colors.yellow : Colors.white,
      child: Text(
        content,
        textAlign: TextAlign.center,
        style: TextStyle(
            fontSize: 16,
            fontWeight: isHeader ? FontWeight.bold : FontWeight.normal),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Queue Simulation')),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: [
            ElevatedButton(
              onPressed: pickAndReadExcelFile,
              child: const Text('Choose and Display Excel File'),
            ),
            const SizedBox(height: 20),
            if (isExcelLoaded)
              Expanded(
                  child:
                      buildTable(excelData, ['code', 'Service', 'Duration'])),
            const SizedBox(height: 20),
            ElevatedButton(
              onPressed: runSimulation,
              child: const Text('Run Simulation'),
            ),
            const SizedBox(height: 20),
            if (simulationData.isNotEmpty)
              Expanded(
                child: buildTable(
                  simulationData,
                  [
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
                  ],
                ),
              ),
            ElevatedButton(
              onPressed: () async {
                service
                    .createExcelTables(); // Call create method to prepare Excel data
                setState(() {
                  isExcelGenerated =
                      true; // Update state to indicate file generation
                });
              },
              child: const Text(
                'Export to Excel',
                style: TextStyle(fontWeight: FontWeight.bold),
              ),
            ),
            ElevatedButton(
              onPressed: isExcelGenerated
                  ? () async {
                      await service.openSavedFile(); // Open the saved file
                    }
                  : null,
              child: const Text(
                'Open Excel File',
                style: TextStyle(fontWeight: FontWeight.bold),
              ),
            ),
          ],
        ),
      ),
    );
  }
}

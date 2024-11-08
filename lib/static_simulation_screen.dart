import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'dart:math';
import 'package:simulation_app/excel_static_service.dart';

class ExcelSimulationScreen extends StatefulWidget {
  const ExcelSimulationScreen({super.key});

  @override
  // ignore: library_private_types_in_public_api
  _ExcelSimulationScreenState createState() => _ExcelSimulationScreenState();
}

class _ExcelSimulationScreenState extends State<ExcelSimulationScreen> {
  List<List<String>> excelData = [];
  List<List<String>> simulationData = [];
  bool isExcelLoaded = false;
  bool isExcelGenerated = false;
  bool isExporting = false;
  String? _filePath = '';
  late ExcelService service;
  int customerNum = 1;
  bool isDarkMode = false; // Track the current theme mode

  final TextEditingController _custNumController = TextEditingController();

  @override
  void initState() {
    super.initState();
    service = ExcelService(_filePath, excelData, simulationData);
  }

  Future<void> pickAndReadExcelFile() async {
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
      );

      if (result != null) {
        _filePath = result.files.single.path;
        var bytes = File(_filePath!).readAsBytesSync();
        var excel = Excel.decodeBytes(bytes);

        if (excel.tables.isNotEmpty) {
          String firstTableName = excel.tables.keys.first;
          var rows = excel.tables[firstTableName]?.rows;

          if (rows != null && rows.isNotEmpty) {
            excelData = rows
                .skip(1)
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
    service = ExcelService(_filePath, excelData, simulationData);
    setState(() {});
  }

  void runSimulation() {
    customerNum = int.tryParse(_custNumController.text) ?? 1;

    if (excelData.isEmpty || excelData.length <= 1) return;

    simulationData.clear();
    int currentTime = 0;
    Random random = Random();

    for (int i = 1; i <= customerNum; i++) {
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

    service = ExcelService(_filePath, excelData, simulationData);
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
                row.add('');
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
      color: isHeader
          ? Colors.yellow
          : (isDarkMode
              ? Colors.grey[850]
              : Theme.of(context).cardColor), // Dark mode color
      child: Text(
        content,
        textAlign: TextAlign.center,
        style: TextStyle(
          fontSize: 16,
          fontWeight: isHeader ? FontWeight.bold : FontWeight.normal,
          color: isHeader
              ? Colors.black
              : (isDarkMode ? Colors.white : null), // Adjust text color
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false, // Remove the debug banner
      theme: isDarkMode
          ? ThemeData.dark()
          : ThemeData.light(), // Set the theme based on the toggle
      home: Scaffold(
        appBar: AppBar(
          title: const Text('Static Simulation'),
          actions: [
            IconButton(
              icon: Icon(
                isDarkMode ? Icons.light_mode : Icons.dark_mode,
              ),
              onPressed: () {
                setState(() {
                  isDarkMode = !isDarkMode; // Toggle the theme mode
                });
              },
            ),
          ],
        ),
        body: SingleChildScrollView(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            children: [
              TextField(
                controller: _custNumController,
                decoration: InputDecoration(
                  labelText: 'Enter Customer Number',
                  border: const OutlineInputBorder(),
                  filled: true,
                  fillColor: Theme.of(context).inputDecorationTheme.fillColor,
                ),
                keyboardType: TextInputType.number,
              ),
              const SizedBox(height: 15),
              Row(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  ElevatedButton(
                    onPressed: pickAndReadExcelFile,
                    style: ElevatedButton.styleFrom(
                      foregroundColor: Colors.white,
                      backgroundColor: Colors.blue,
                      shadowColor: Colors.black,
                      shape: const RoundedRectangleBorder(
                        borderRadius: BorderRadius.all(Radius.circular(5)),
                      ),
                    ),
                    child: const Text('Choose Excel File'),
                  ),
                  const SizedBox(width: 15),
                  ElevatedButton(
                    onPressed: runSimulation,
                    style: ElevatedButton.styleFrom(
                      foregroundColor: Colors.white,
                      backgroundColor: Colors.green,
                      shadowColor: Colors.black,
                      shape: const RoundedRectangleBorder(
                        borderRadius: BorderRadius.all(Radius.circular(5)),
                      ),
                    ),
                    child: const Text('Run Simulation'),
                  ),
                ],
              ),
              const SizedBox(height: 25),
              if (isExcelLoaded)
                buildTable(excelData, ['code', 'Service', 'Duration']),
              const SizedBox(height: 30),
              if (simulationData.isNotEmpty)
                buildTable(
                  simulationData,
                  [
                    'Cust_id',
                    'Interval',
                    'Arr.Clock',
                    'code',
                    'Service',
                    'Start',
                    'Duration',
                    'EndClock',
                    'Serv.State',
                    'Cust Wait'
                  ],
                ),
              const SizedBox(height: 20),
              if (simulationData.isNotEmpty)
                Stack(
                  children: [
                    Row(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                        ElevatedButton(
                          onPressed: isExporting
                              ? null // Disable button while exporting
                              : () async {
                                  setState(() {
                                    isExporting = true;
                                  });

                                  await service.createExcelTables();

                                  setState(() {
                                    isExporting = false;
                                    isExcelGenerated = true;
                                  });
                                },
                          style: ElevatedButton.styleFrom(
                            foregroundColor: Colors.white,
                            backgroundColor: Colors.teal,
                            shadowColor: Colors.black,
                            shape: const RoundedRectangleBorder(
                              borderRadius:
                                  BorderRadius.all(Radius.circular(5)),
                            ),
                          ),
                          child: isExporting
                              ? const SizedBox(
                                  height:
                                      24, // Set the height of the circular indicator
                                  width:
                                      24, // Set the width of the circular indicator
                                  child: CircularProgressIndicator(
                                    color: Colors
                                        .white, // Adjust the indicator color to fit the theme
                                    strokeWidth:
                                        3, // Set the stroke width for a more refined look
                                  ),
                                )
                              : const Text(
                                  'Export to Excel',
                                  style: TextStyle(
                                    fontWeight: FontWeight.bold,
                                    fontSize:
                                        16, // Ensure the font size is appropriate for readability
                                    color: Colors
                                        .white, // Text color for light contrast
                                  ),
                                ),
                        ),
                        const SizedBox(width: 15),
                        ElevatedButton(
                          onPressed: isExcelGenerated
                              ? () async {
                                  await service.openSavedFile();
                                }
                              : null,
                          style: ElevatedButton.styleFrom(
                            foregroundColor: Colors.white,
                            backgroundColor: Colors.red,
                            shadowColor: Colors.black,
                            shape: const RoundedRectangleBorder(
                              borderRadius:
                                  BorderRadius.all(Radius.circular(5)),
                            ),
                          ),
                          child: const Text('Open Excel File'),
                        ),
                      ],
                    ),
                  ],
                ),
            ],
          ),
        ),
      ),
    );
  }
}

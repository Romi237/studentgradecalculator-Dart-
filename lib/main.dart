import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'dart:typed_data';
import 'package:pdf/widgets.dart' as pw;
import 'package:pdf/pdf.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';
import 'dart:io' show Platform, File;
import 'dart:html' show Blob, AnchorElement, Url; // Fixed import syntax

void main() => runApp(const GradeCalculatorApp());

class GradeCalculatorApp extends StatelessWidget {
  const GradeCalculatorApp({super.key});

  @override
  Widget build(BuildContext context) => MaterialApp(
        title: 'Student Grade Calculator',
        theme: ThemeData(primarySwatch: Colors.blue, useMaterial3: true),
        home: const GradeCalculatorHome(),
      );
}

class GradeCalculatorHome extends StatefulWidget {
  const GradeCalculatorHome({super.key});

  @override
  State<GradeCalculatorHome> createState() => _GradeCalculatorHomeState();
}

class _GradeCalculatorHomeState extends State<GradeCalculatorHome> {
  List<Student> students = [];
  bool isLoading = false;
  String? selectedFile;
  final _calculator = GradeCalculator();

  @override
  Widget build(BuildContext context) => Scaffold(
        appBar: AppBar(title: const Text('Student Grade Calculator')),
        body: Padding(
          padding: const EdgeInsets.all(20),
          child: Column(children: [
            _buildFileCard(),
            const SizedBox(height: 20),
            _buildProcessButton(),
            const SizedBox(height: 20),
            _buildResultsCard(),
            if (students.isNotEmpty) _buildExportButtons(),
            if (isLoading) const LinearProgressIndicator(),
          ]),
        ),
      );

  Widget _buildFileCard() => Card(
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              const Text('Input File',
                  style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 10),
              Row(children: [
                Flexible(
                  child: Text(
                    selectedFile ?? 'No file selected',
                    overflow: TextOverflow.ellipsis,
                  ),
                ),
                const SizedBox(width: 10),
                ElevatedButton.icon(
                  onPressed: isLoading ? null : _pickFile,
                  icon: const Icon(Icons.file_upload),
                  label: const Text('Select Excel'),
                ),
              ]),
            ],
          ),
        ),
      );

  Widget _buildProcessButton() => Center(
        child: ElevatedButton.icon(
          onPressed: students.isEmpty || isLoading ? null : _processGrades,
          icon: const Icon(Icons.calculate),
          label: const Text('Calculate Grades'),
          style: ElevatedButton.styleFrom(
              padding:
                  const EdgeInsets.symmetric(horizontal: 40, vertical: 15)),
        ),
      );

  Widget _buildResultsCard() => Expanded(
        child: Card(
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Text('Results',
                        style: TextStyle(
                            fontSize: 18, fontWeight: FontWeight.bold)),
                    if (students.isNotEmpty)
                      Text('${students.length} students'),
                  ],
                ),
                const SizedBox(height: 10),
                Expanded(
                  child: students.isEmpty
                      ? const Center(
                          child: Text('No data loaded. Please select a file.'))
                      : ListView.builder(
                          itemCount: students.length,
                          itemBuilder: (_, i) => Card(
                            margin: const EdgeInsets.only(bottom: 8),
                            child: ListTile(
                              leading: CircleAvatar(
                                backgroundColor:
                                    _getGradeColor(students[i].letterGrade),
                                child: Text(students[i].letterGrade),
                              ),
                              title: Text(students[i].displayName),
                              subtitle: Text(
                                  'Score: ${students[i].score?.toStringAsFixed(2) ?? 'N/A'}'),
                            ),
                          ),
                        ),
                ),
              ],
            ),
          ),
        ),
      );

  Widget _buildExportButtons() => Row(
        mainAxisAlignment: MainAxisAlignment.end,
        children: [
          TextButton.icon(
            onPressed: isLoading ? null : _exportToExcel,
            icon: const Icon(Icons.table_chart),
            label: const Text('Export Excel'),
          ),
          const SizedBox(width: 10),
          ElevatedButton.icon(
            onPressed: isLoading ? null : _exportToPDF,
            icon: const Icon(Icons.picture_as_pdf),
            label: const Text('Export PDF'),
          ),
        ],
      );

  Color _getGradeColor(String g) => switch (g) {
        'A' => Colors.green,
        'B' => Colors.lightGreen,
        'C' => Colors.orange,
        'D' => Colors.deepOrange,
        'F' => Colors.red,
        _ => Colors.grey,
      };

  Future<void> _pickFile() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
    );
    if (result == null) return;

    setState(() {
      selectedFile = result.files.single.name;
      isLoading = true;
      students = [];
    });

    await _parseFile(result.files.single);
    setState(() => isLoading = false);
  }

  Future<void> _parseFile(PlatformFile file) async {
    try {
      if (!(file.extension?.contains('xls') ?? false)) {
        return _showDialog('Error', 'Please select an Excel file');
      }

      if (file.bytes == null) {
        return _showDialog('Error', 'Could not read file bytes');
      }

      final excel = Excel.decodeBytes(file.bytes!);
      final List<Student> parsed = [];

      for (var table in excel.tables.keys) {
        var sheet = excel.tables[table];
        if (sheet == null) continue;

        for (int i = 1; i < sheet.rows.length; i++) {
          var row = sheet.rows[i];
          if (row.length < 2) continue;

          var name = row[0]?.value?.toString().trim();
          var score = double.tryParse(row[1]?.value?.toString() ?? '');

          if (name != null && name.isNotEmpty) {
            parsed.add(Student(name: name, score: score));
          }
        }
      }

      setState(() => students = parsed);
      _showDialog('Success', 'Loaded ${parsed.length} students');
    } catch (e) {
      _showDialog('Error', 'Failed to parse file: $e');
    }
  }

  void _processGrades() {
    var stats = _calculator.calculateStatistics(students);
    var passed =
        students.where((s) => s.score != null && s.score! >= 60).length;
    var avg = students
            .where((s) => s.score != null)
            .map((s) => s.score!)
            .fold(0.0, (a, b) => a + b) /
        students.where((s) => s.score != null).length;

    showDialog(
      context: context,
      builder: (_) => AlertDialog(
        title: const Text('Statistics'),
        content: SingleChildScrollView(
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            Text('Total: ${stats['total']}'),
            Text('Average: ${avg.toStringAsFixed(2)}'),
            Text('Passed: $passed'),
            Text('Failed: ${stats['total'] - passed}'),
            const Divider(),
            ...stats.entries
                .where((e) => e.key != 'total')
                .map((e) => Text('${e.key}: ${e.value}')),
          ]),
        ),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(context), child: const Text('OK'))
        ],
      ),
    );
  }

  Future<void> _exportToExcel() async {
    var excel = Excel.createExcel();
    var sheet = excel['Grades'];
    sheet.appendRow(['Name', 'Score', 'Grade']);
    students.forEach((s) => sheet.appendRow(
        [s.displayName, s.score?.toStringAsFixed(2) ?? 'N/A', s.letterGrade]));

    final bytes = excel.save();
    if (bytes == null) return;

    final timestamp = DateTime.now().millisecondsSinceEpoch;

    if (Platform.isWindows ||
        Platform.isLinux ||
        Platform.isMacOS ||
        Platform.isAndroid ||
        Platform.isIOS) {
      var dir = await getApplicationDocumentsDirectory();
      var path = '${dir.path}/grades_$timestamp.xlsx';
      File(path).writeAsBytesSync(bytes);
      _showDialog('Success', 'Saved to: $path');
      OpenFile.open(path);
    } else {
      final blob = Blob([bytes]);
      final url = Url.createObjectUrlFromBlob(blob);
      AnchorElement(href: url)
        ..download = 'grades_$timestamp.xlsx'
        ..click();
      Url.revokeObjectUrl(url);
      _showDialog('Success', 'File downloaded');
    }
  }

  Future<void> _exportToPDF() async {
    final pdf = pw.Document();
    pdf.addPage(pw.MultiPage(
      pageFormat: PdfPageFormat.a4,
      build: (_) => [
        pw.Header(level: 0, child: pw.Text('Student Grades')),
        pw.TableHelper.fromTextArray(
          headers: ['Name', 'Score', 'Grade'],
          data: students
              .map((s) => [
                    s.displayName,
                    s.score?.toStringAsFixed(2) ?? 'N/A',
                    s.letterGrade
                  ])
              .toList(),
        ),
      ],
    ));

    final bytes = await pdf.save();
    final timestamp = DateTime.now().millisecondsSinceEpoch;

    if (Platform.isWindows ||
        Platform.isLinux ||
        Platform.isMacOS ||
        Platform.isAndroid ||
        Platform.isIOS) {
      var dir = await getApplicationDocumentsDirectory();
      var path = '${dir.path}/grades_$timestamp.pdf';
      await File(path).writeAsBytes(bytes);
      _showDialog('Success', 'Saved to: $path');
      OpenFile.open(path);
    } else {
      final blob = Blob([bytes]);
      final url = Url.createObjectUrlFromBlob(blob);
      AnchorElement(href: url)
        ..download = 'grades_$timestamp.pdf'
        ..click();
      Url.revokeObjectUrl(url);
      _showDialog('Success', 'File downloaded');
    }
  }

  void _showDialog(String title, String msg) => showDialog(
        context: context,
        builder: (_) => AlertDialog(
          title: Text(title),
          content: SingleChildScrollView(child: Text(msg)),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('OK'))
          ],
        ),
      );
}

class Student {
  final String name;
  final double? score;
  Student({required this.name, this.score});

  String get displayName => name.trim().isEmpty ? 'Unknown' : name.trim();
  String get letterGrade => score == null
      ? 'N/A'
      : score! < 0 || score! > 100
          ? 'Invalid'
          : switch (score!.toInt()) {
              >= 90 => 'A',
              >= 80 => 'B',
              >= 70 => 'C',
              >= 60 => 'D',
              _ => 'F'
            };
}

class GradeCalculator {
  Map<String, dynamic> calculateStatistics(List<Student> students) {
    var counts = students
        .map((s) => s.letterGrade)
        .fold<Map<String, int>>({}, (m, g) => m..[g] = (m[g] ?? 0) + 1);
    counts['total'] = students.length;
    return counts;
  }
}

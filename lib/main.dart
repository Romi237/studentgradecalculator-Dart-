import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'dart:typed_data';
import 'package:pdf/widgets.dart' as pw;
import 'package:pdf/pdf.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';

class Student {
  final String name;
  final double? score;

  Student({required this.name, this.score});

  /// Human‑readable name; fall back to a default when the caller provided an
  /// empty string or only whitespace.
  String get displayName {
    final trimmed = name.trim();
    return trimmed.isEmpty ? 'Unknown Student' : trimmed;
  }

  /// Convert the numeric score into a letter grade.  The method is tolerant
  /// of `null` and out‑of‑range values so callers don't need to perform
  /// additional checks.
  String get letterGrade {
    if (score == null) return 'N/A';
    if (score! < 0 || score! > 100) return 'Invalid';

    switch (score!.toInt()) {
      case 100:
      case 99:
      case 98:
      case 97:
      case 96:
      case 95:
      case 94:
      case 93:
      case 92:
      case 91:
      case 90:
        return 'A';
      case 89:
      case 88:
      case 87:
      case 86:
      case 85:
      case 84:
      case 83:
      case 82:
      case 81:
      case 80:
        return 'B';
      case 79:
      case 78:
      case 77:
      case 76:
      case 75:
      case 74:
      case 73:
      case 72:
      case 71:
      case 70:
        return 'C';
      case 69:
      case 68:
      case 67:
      case 66:
      case 65:
      case 64:
      case 63:
      case 62:
      case 61:
      case 60:
        return 'D';
      default:
        return 'F';
    }
  }
}

/// Prompt the user repeatedly until they enter a valid score between 0 and
/// 100.  Returns `null` only if the user hits EOF (Ctrl‑Z on Windows or
/// Ctrl‑D on *nix).
double? promptForScore() {
  while (true) {
    stdout.write('Enter student score (0–100): ');
    final line = stdin.readLineSync();
    if (line == null) return null; // EOF

    final value = double.tryParse(line);
    if (value == null) {
      print('Invalid input – please type a number.');
      continue;
    }
    if (value < 0 || value > 100) {
      print('Score must be between 0 and 100.');
      continue;
    }
    return value;
  }
}

/// The entry point for the console application.  Prompts the user for a name
/// and score, constructs a [Student], and prints out the derived grade.
void main() {
  print('=== Student Grade Calculator ===');

  stdout.write('Enter student name: ');
  final rawName = stdin.readLineSync();
  final name = rawName ?? ''; // null → empty string for simplicity

  final score = promptForScore();

  final student = Student(name: name, score: score);

  print('\n--- Result ---');
  print('Name : ${student.displayName}');
  print('Score: ${student.score?.toStringAsFixed(2) ?? 'N/A'}');
  print('Grade: ${student.letterGrade}');
}

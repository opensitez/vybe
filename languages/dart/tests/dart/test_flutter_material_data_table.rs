use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material DataTable
// ═══════════════════════════════════════════════════════════

#[test]
fn data_table_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dt = DataTable(columns: const [], rows: const []);
  print(dt is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn data_table_columns_rows() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dt = DataTable(
    columns: const [
      DataColumn(label: Text('A')),
      DataColumn(label: Text('B')),
    ],
    rows: const [
      DataRow(cells: [DataCell(Text('1')), DataCell(Text('2'))]),
    ],
  );
  print('${dt.columns.length}:${dt.rows.length}');
}
"#
        ),
        vec!["2:1"]
    );
}

#[test]
fn data_table_sort_column_index() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  final dt = DataTable(
    sortColumnIndex: 0,
    sortAscending: true,
    columns: const [],
    rows: const [],
  );
  print('${dt.sortColumnIndex}:${dt.sortAscending}');
}
"#
        ),
        vec!["0:true"]
    );
}

#[test]
fn data_column_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const dc = DataColumn(
    label: Text('ID'),
    tooltip: 'Identifier',
    numeric: true,
  );
  print('${dc.tooltip}:${dc.numeric}');
}
"#
        ),
        vec!["Identifier:true"]
    );
}

#[test]
fn data_row_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const dr = DataRow(
    selected: true,
    cells: [],
  );
  print(dr.selected);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn data_cell_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const dc = DataCell(Text('Value'), showEditIcon: true);
  print(dc.showEditIcon);
}
"#
        ),
        vec!["true"]
    );
}

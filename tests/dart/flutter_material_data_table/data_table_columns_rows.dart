// vybe-test: dart/flutter_material_data_table/data_table_columns_rows
// origin: languages/dart/tests/dart/test_flutter_material_data_table.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

import 'package:flutter/material.dart';
void __vybeMain() {
  final dt = DataTable(
    columns: const [
      DataColumn(label: Text('A')),
      DataColumn(label: Text('B')),
    ],
    rows: const [
      DataRow(cells: [DataCell(Text('1')), DataCell(Text('2'))]),
    ],
  );
  __p('${dt.columns.length}:${dt.rows.length}');
}

void main() {
  __vybeMain();
  __check('2:1');
}

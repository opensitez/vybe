// vybe-test: dart/flutter_material_grid_tile/grid_tile_bar_properties
// origin: languages/dart/tests/dart/test_flutter_material_grid_tile.rs

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
  const gtb = GridTileBar(
    backgroundColor: Color(0xFF00FF00),
    leading: Icon(Icons.star),
    title: Text('Title'),
    subtitle: Text('Sub'),
    trailing: Icon(Icons.more_vert),
  );
  __p('${gtb.backgroundColor?.value == 0xFF00FF00}:${gtb.leading is Icon}:${gtb.title is Text}:${gtb.subtitle is Text}:${gtb.trailing is Icon}');
}

void main() {
  __vybeMain();
  __check('true:true:true:true:true');
}

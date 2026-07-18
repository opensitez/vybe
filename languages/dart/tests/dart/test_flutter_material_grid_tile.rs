use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: material GridTile
// ═══════════════════════════════════════════════════════════

#[test]
fn grid_tile_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const gt = GridTile(child: SizedBox());
  print(gt is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_tile_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const gt = GridTile(child: Placeholder());
  print(gt.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_tile_header() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const gt = GridTile(
    header: Text('Header'),
    child: SizedBox(),
  );
  print(gt.header is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_tile_footer() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const gt = GridTile(
    footer: Text('Footer'),
    child: SizedBox(),
  );
  print(gt.footer is Text);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_tile_bar_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const gtb = GridTileBar(title: Text('Title'));
  print(gtb is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn grid_tile_bar_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/material.dart';
void main() {
  const gtb = GridTileBar(
    backgroundColor: Color(0xFF00FF00),
    leading: Icon(Icons.star),
    title: Text('Title'),
    subtitle: Text('Sub'),
    trailing: Icon(Icons.more_vert),
  );
  print('${gtb.backgroundColor?.value == 0xFF00FF00}:${gtb.leading is Icon}:${gtb.title is Text}:${gtb.subtitle is Text}:${gtb.trailing is Icon}');
}
"#
        ),
        vec!["true:true:true:true:true"]
    );
}

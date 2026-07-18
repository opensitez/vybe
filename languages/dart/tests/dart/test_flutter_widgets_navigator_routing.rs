use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Navigator & Routing
// ═══════════════════════════════════════════════════════════

#[test]
fn route_settings_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final settings = RouteSettings(name: '/home', arguments: 'data');
  print('${settings.name}:${settings.arguments}');
}
"#
        ),
        vec!["/home:data"]
    );
}

#[test]
fn route_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyRoute extends Route<String> {
  @override
  bool get maintainState => true;
}
void main() {
  final r = MyRoute();
  print(r != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn navigator_observer() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyObserver extends NavigatorObserver {
  @override
  void didPush(Route<dynamic> route, Route<dynamic>? previousRoute) {
    print('pushed');
  }
}
void main() {
  final o = MyObserver();
  print(o != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn page_route_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder(
    pageBuilder: (context, anim, secondAnim) => const SizedBox(),
    transitionDuration: Duration(milliseconds: 500),
  );
  print(r.transitionDuration.inMilliseconds);
}
"#
        ),
        vec!["500"]
    );
}

#[test]
fn navigator_push_abstract() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  // It's hard to mock a full Navigator tree headless,
  // we just test API compilation for Route class methods
  final r = PageRouteBuilder(pageBuilder: (c, a, sa) => const SizedBox());
  print(r.settings.name == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn route_popped_future() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder<int>(pageBuilder: (c, a, sa) => const SizedBox());
  print(r.popped is Future);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn route_did_pop() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder<int>(pageBuilder: (c, a, sa) => const SizedBox());
  print(r.didPop(1));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn route_is_first() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder<int>(pageBuilder: (c, a, sa) => const SizedBox());
  print(r.isFirst);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn route_will_pop() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() async {
  final r = PageRouteBuilder<int>(pageBuilder: (c, a, sa) => const SizedBox());
  final res = await r.willPop();
  print(res == RoutePopDisposition.pop);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn route_active_status() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = PageRouteBuilder<int>(pageBuilder: (c, a, sa) => const SizedBox());
  print(r.isActive);
}
"#
        ),
        vec!["false"]
    );
}

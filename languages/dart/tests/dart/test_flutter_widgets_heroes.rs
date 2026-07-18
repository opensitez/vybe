use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Heroes
// ═══════════════════════════════════════════════════════════

#[test]
fn hero_widget_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final h = Hero(tag: 'my_tag', child: const SizedBox());
  print(h.tag);
}
"#
        ),
        vec!["my_tag"]
    );
}

#[test]
fn hero_widget_transition_on_user_gestures() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final h = Hero(tag: 'tag', child: const SizedBox(), transitionOnUserGestures: true);
  print(h.transitionOnUserGestures);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn hero_widget_flight_shuttle_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  Widget builder(BuildContext flightContext, Animation<double> animation, HeroFlightDirection flightDirection, BuildContext fromHeroContext, BuildContext toHeroContext) {
    return const SizedBox();
  }
  final h = Hero(tag: 'tag', child: const SizedBox(), flightShuttleBuilder: builder);
  print(h.flightShuttleBuilder != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn hero_widget_placeholder_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  Widget builder(BuildContext context, Size heroSize, Widget child) {
    return const SizedBox();
  }
  final h = Hero(tag: 'tag', child: const SizedBox(), placeholderBuilder: builder);
  print(h.placeholderBuilder != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn hero_controller_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final hc = HeroController();
  print(hc != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn hero_mode_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final hm = HeroMode(enabled: false, child: const SizedBox());
  print(hm.enabled);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn hero_flight_direction_values() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  print(HeroFlightDirection.push.name);
  print(HeroFlightDirection.pop.name);
}
"#
        ),
        vec!["push\npop"]
    );
}

#[test]
fn hero_tag_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final h1 = Hero(tag: 123, child: const SizedBox());
  final h2 = Hero(tag: 'abc', child: const SizedBox());
  print(h1.tag is int);
  print(h2.tag is String);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn hero_element_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final h = Hero(tag: 'test', child: const SizedBox());
  final e = h.createElement();
  print(e is StatefulWidget); // Hero is usually a StatefulWidget
}
"#
        ),
        // Hero extends StatefulWidget, so e is StatefulElement
        vec!["false"] // e is Element, not StatefulWidget
    );
}

#[test]
fn hero_element_is_stateful() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final h = Hero(tag: 'test', child: const SizedBox());
  final e = h.createElement();
  print(e is StatefulElement);
}
"#
        ),
        vec!["true"]
    );
}

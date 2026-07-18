use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SliverList
// ═══════════════════════════════════════════════════════════

#[test]
fn sliver_list_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sl = SliverList(
    delegate: SliverChildListDelegate(const []),
  );
  print(sl is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_list_builder() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sl = SliverList.builder(
    itemCount: 10,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(sl is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_list_separated() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sl = SliverList.separated(
    itemCount: 5,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
    separatorBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  print(sl is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_list_list_delegate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sl = SliverList(
    delegate: SliverChildListDelegate(const []),
  );
  print(sl.delegate is SliverChildDelegate);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sliver_list_builder_delegate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sl = SliverList(
    delegate: SliverChildBuilderDelegate((BuildContext context, int index) => const SizedBox(), childCount: 0),
  );
  print(sl.delegate is SliverChildBuilderDelegate);
}
"#
        ),
        vec!["true"]
    );
}

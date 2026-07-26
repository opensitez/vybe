use super::helpers::run_prints;

#[test]
fn test_property_hooks_field_backing_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
class CounterDemo {
    public int $count = 0 {
        set => $field + $value;
    }
}
$c = new CounterDemo();
$c->count = 5;
$c->count = 10;
echo $c->count, "\n";
"#
        ),
        vec!["15"]
    );
}

#[test]
fn test_property_hooks_backing_with_validation() {
    assert_eq!(
        run_prints(
            r#"<?php
class Score {
    public int $value = 0 {
        set => max(0, $value);
    }
}
$s = new Score();
$s->value = -3;
echo $s->value;
"#,
        ),
        vec!["0"]
    );
}

#[test]
fn test_property_hooks_field_backing_preserves_previous_when_no_set() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public int $count = 1 {
        set => $field;
    }
}
$c = new Counter();
echo $c->count;
"#,
        ),
        vec!["1"]
    );
}

#[test]
fn test_property_hooks_backing_set_defaulted_when_ignored_input() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public int $count = 2 {
        set => $field + 1;
    }
}
$c = new Counter();
$c->count = 4;
$c->count = 0;
echo $c->count;
"#,
        ),
        vec!["3"]
    );
}

#[test]
fn test_property_hooks_backing_set_with_clamp_and_validation() {
    assert_eq!(
        run_prints(
            r#"<?php
class Grade {
    public int $value = 10 {
        set => max(0, min(100, $value));
    }
}
$g = new Grade();
$g->value = 120;
echo $g->value . '|';
$g->value = -10;
echo $g->value;
"#,
        ),
        vec!["100|0"]
    );
}

#[test]
fn test_property_hooks_backing_get_uses_virtual_projection() {
    assert_eq!(
        run_prints(
            r#"<?php
class Price {
    public int $cents = 1250 {
        get => $field / 100;
        set => (int) round($value * 100);
    }
}
$p = new Price();
$p->cents = 12.5;
echo $p->cents;
"#,
        ),
        vec!["12.5"]
    );
}

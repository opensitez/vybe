//! `volatile` field reads and writes — structural visibility tests via printed counts.
//! GAP: concurrency primitives.

use crate::csharp_cases;

csharp_cases! {
    volatile_int_write_then_read_count => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
box.Value = 7;
Console.WriteLine(box.Value);
"#,
        ["7"]
    };

    volatile_int_default_zero_count => {
        r#"
class FlagBox {
    public volatile int Value;
}
var box = new FlagBox();
Console.WriteLine(box.Value);
"#,
        ["0"]
    };

    volatile_int_increment_via_local_count => {
        r#"
class FlagBox {
    public volatile int Value = 1;
}
var box = new FlagBox();
int snapshot = box.Value;
box.Value = snapshot + 2;
Console.WriteLine(box.Value);
"#,
        ["3"]
    };

    volatile_bool_true_read_count => {
        r#"
class FlagBox {
    public volatile bool Ready = true;
}
var box = new FlagBox();
Console.WriteLine(box.Ready ? 1 : 0);
"#,
        ["1"]
    };

    volatile_bool_false_read_count => {
        r#"
class FlagBox {
    public volatile bool Ready = false;
}
var box = new FlagBox();
Console.WriteLine(box.Ready ? 1 : 0);
"#,
        ["0"]
    };

    volatile_bool_toggle_count => {
        r#"
class FlagBox {
    public volatile bool Ready = false;
}
var box = new FlagBox();
box.Ready = true;
Console.WriteLine(box.Ready ? 1 : 0);
"#,
        ["1"]
    };

    volatile_long_write_read_count => {
        r#"
class FlagBox {
    public volatile long Value = 0L;
}
var box = new FlagBox();
box.Value = 1000000L;
Console.WriteLine(box.Value);
"#,
        ["1000000"]
    };

    volatile_static_int_field_count => {
        r#"
class FlagBox {
    public static volatile int Shared = 0;
}
FlagBox.Shared = 12;
Console.WriteLine(FlagBox.Shared);
"#,
        ["12"]
    };

    volatile_two_fields_independent_counts => {
        r#"
class FlagBox {
    public volatile int A = 1;
    public volatile int B = 2;
}
var box = new FlagBox();
Console.WriteLine(box.A + box.B);
"#,
        ["3"]
    };

    volatile_read_after_multiple_writes_last_wins => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
box.Value = 1;
box.Value = 2;
box.Value = 9;
Console.WriteLine(box.Value);
"#,
        ["9"]
    };

    volatile_int_negative_value_count => {
        r#"
class FlagBox {
    public volatile int Value = -4;
}
var box = new FlagBox();
Console.WriteLine(box.Value);
"#,
        ["-4"]
    };

    volatile_int_assign_from_expression => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
box.Value = 3 + 4;
Console.WriteLine(box.Value);
"#,
        ["7"]
    };

    volatile_bool_assign_from_comparison => {
        r#"
class FlagBox {
    public volatile bool Ready = false;
}
var box = new FlagBox();
box.Ready = 5 > 3;
Console.WriteLine(box.Ready ? 1 : 0);
"#,
        ["1"]
    };

    volatile_field_in_method_read_count => {
        r#"
class FlagBox {
    public volatile int Value = 6;
    public int Read() { return Value; }
}
var box = new FlagBox();
Console.WriteLine(box.Read());
"#,
        ["6"]
    };

    volatile_field_in_method_write_count => {
        r#"
class FlagBox {
    public volatile int Value = 0;
    public void Write(int n) { Value = n; }
}
var box = new FlagBox();
box.Write(15);
Console.WriteLine(box.Value);
"#,
        ["15"]
    };

    volatile_loop_write_accumulator_count => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
for (int i = 1; i <= 4; i++) box.Value += i;
Console.WriteLine(box.Value);
"#,
        ["10"]
    };

    volatile_copy_to_local_preserves_count => {
        r#"
class FlagBox {
    public volatile int Value = 8;
}
var box = new FlagBox();
int local = box.Value;
Console.WriteLine(local);
"#,
        ["8"]
    };

    volatile_static_bool_flag_count => {
        r#"
class FlagBox {
    public static volatile bool Done = false;
}
FlagBox.Done = true;
Console.WriteLine(FlagBox.Done ? 1 : 0);
"#,
        ["1"]
    };

    volatile_int_if_branch_read_count => {
        r#"
class FlagBox {
    public volatile int Value = 2;
}
var box = new FlagBox();
int count = 0;
if (box.Value == 2) count = 1;
Console.WriteLine(count);
"#,
        ["1"]
    };

    volatile_int_switch_read_count => {
        r#"
class FlagBox {
    public volatile int Value = 3;
}
var box = new FlagBox();
int count = 0;
switch (box.Value) {
    case 3: count = 30; break;
    default: count = 0; break;
}
Console.WriteLine(count);
"#,
        ["30"]
    };

    volatile_bool_ternary_read_count => {
        r#"
class FlagBox {
    public volatile bool Ready = true;
}
var box = new FlagBox();
int count = box.Ready ? 5 : 0;
Console.WriteLine(count);
"#,
        ["5"]
    };

    volatile_two_instances_isolated_counts => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var a = new FlagBox();
var b = new FlagBox();
a.Value = 4;
b.Value = 5;
Console.WriteLine(a.Value + b.Value);
"#,
        ["9"]
    };

    volatile_read_in_expression_sum => {
        r#"
class FlagBox {
    public volatile int X = 2;
    public volatile int Y = 3;
}
var box = new FlagBox();
Console.WriteLine(box.X + box.Y);
"#,
        ["5"]
    };

    volatile_write_from_parameter_count => {
        r#"
class FlagBox {
    public volatile int Value = 0;
    public void Set(int n) { Value = n; }
}
var box = new FlagBox();
box.Set(22);
Console.WriteLine(box.Value);
"#,
        ["22"]
    };

    volatile_int_max_value_write => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
box.Value = 2147483647;
Console.WriteLine(box.Value > 0 ? 1 : 0);
"#,
        ["1"]
    };

    volatile_int_zero_write_count => {
        r#"
class FlagBox {
    public volatile int Value = 99;
}
var box = new FlagBox();
box.Value = 0;
Console.WriteLine(box.Value);
"#,
        ["0"]
    };

    volatile_bool_not_operator_count => {
        r#"
class FlagBox {
    public volatile bool Ready = false;
}
var box = new FlagBox();
Console.WriteLine(!box.Ready ? 1 : 0);
"#,
        ["1"]
    };

    volatile_field_passed_to_method_by_value => {
        r#"
int Double(int n) { return n * 2; }
class FlagBox {
    public volatile int Value = 5;
}
var box = new FlagBox();
Console.WriteLine(Double(box.Value));
"#,
        ["10"]
    };

    volatile_read_while_loop_count => {
        r#"
class FlagBox {
    public volatile int Value = 3;
}
var box = new FlagBox();
int count = 0;
while (box.Value > 0) {
    count++;
    box.Value--;
}
Console.WriteLine(count);
"#,
        ["3"]
    };

    volatile_do_while_read_once_count => {
        r#"
class FlagBox {
    public volatile int Value = 1;
}
var box = new FlagBox();
int count = 0;
do {
    count += box.Value;
    box.Value = 0;
} while (box.Value > 0);
Console.WriteLine(count);
"#,
        ["1"]
    };

    volatile_static_increment_via_local_count => {
        r#"
class FlagBox {
    public static volatile int Shared = 10;
}
int snap = FlagBox.Shared;
FlagBox.Shared = snap + 1;
Console.WriteLine(FlagBox.Shared);
"#,
        ["11"]
    };

    volatile_long_negative_read => {
        r#"
class FlagBox {
    public volatile long Value = -500L;
}
var box = new FlagBox();
Console.WriteLine(box.Value);
"#,
        ["-500"]
    };

    volatile_bool_and_expression_count => {
        r#"
class FlagBox {
    public volatile bool A = true;
    public volatile bool B = true;
}
var box = new FlagBox();
Console.WriteLine((box.A && box.B) ? 1 : 0);
"#,
        ["1"]
    };

    volatile_bool_or_expression_count => {
        r#"
class FlagBox {
    public volatile bool A = false;
    public volatile bool B = true;
}
var box = new FlagBox();
Console.WriteLine((box.A || box.B) ? 1 : 0);
"#,
        ["1"]
    };

    volatile_int_post_read_assign_count => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
box.Value = 1;
int first = box.Value;
box.Value = 2;
int second = box.Value;
Console.WriteLine(first + second);
"#,
        ["3"]
    };

    volatile_nested_class_field_count => {
        r#"
class Outer {
    public class Inner {
        public volatile int Value = 0;
    }
}
var inner = new Outer.Inner();
inner.Value = 13;
Console.WriteLine(inner.Value);
"#,
        ["13"]
    };

    volatile_read_after_constructor_set => {
        r#"
class FlagBox {
    public volatile int Value;
    public FlagBox(int n) { Value = n; }
}
var box = new FlagBox(18);
Console.WriteLine(box.Value);
"#,
        ["18"]
    };

    volatile_task_run_write_read_count => {
        r#"
class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
System.Threading.Tasks.Task.Run(() => { box.Value = 6; }).Wait();
Console.WriteLine(box.Value);
"#,
        ["6"]
    };

    volatile_multiple_reads_same_value_count => {
        r#"
class FlagBox {
    public volatile int Value = 4;
}
var box = new FlagBox();
int count = box.Value + box.Value + box.Value;
Console.WriteLine(count);
"#,
        ["12"]
    };

    volatile_bool_equality_check_count => {
        r#"
class FlagBox {
    public volatile bool Ready = true;
}
var box = new FlagBox();
Console.WriteLine(box.Ready == true ? 1 : 0);
"#,
        ["1"]
    };
}

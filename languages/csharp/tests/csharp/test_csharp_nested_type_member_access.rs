//! Nested types reach outer private members per C# accessibility rules.
use super::helpers::run_csharp;

#[test]
fn nested_class_can_read_outer_private_instance_field() {
    assert_eq!(
        run_csharp(
            r#"
class Outer {
    int secret = 8;
    class Inner {
        Outer parent;
        public Inner(Outer parent) { this.parent = parent; }
        public int Read() { return parent.secret; }
    }
    public int ViaInner() { return new Inner(this).Read(); }
}
Console.WriteLine(new Outer().ViaInner());
"#
        ),
        &["8"]
    );
}

#[test]
fn nested_class_can_invoke_outer_private_instance_method() {
    assert_eq!(
        run_csharp(
            r#"
class Outer {
    int Twice(int n) { return n * 2; }
    class Inner {
        Outer parent;
        public Inner(Outer parent) { this.parent = parent; }
        public int Run(int n) { return parent.Twice(n); }
    }
    public int ViaInner(int n) { return new Inner(this).Run(n); }
}
Console.WriteLine(new Outer().ViaInner(5));
"#
        ),
        &["10"]
    );
}

#[test]
fn nested_static_class_reads_outer_static_private_state() {
    assert_eq!(
        run_csharp(
            r#"
class Outer {
    static int tally = 3;
    static class Inner {
        public static int Read() { return tally; }
    }
    public static int Via() { return Inner.Read(); }
}
Console.WriteLine(Outer.Via());
"#
        ),
        &["3"]
    );
}

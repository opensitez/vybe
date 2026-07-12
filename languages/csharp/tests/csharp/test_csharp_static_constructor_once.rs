//! Static constructor (type initializer) runs once before static members are used.
use super::helpers::run_csharp;

#[test]
fn static_constructor_increments_counter_only_once_across_two_instance_allocations() {
    assert_eq!(
        run_csharp(
            r#"
class Tracker {
    public static int Instances;
    static Tracker() { Instances++; }
}
_ = new Tracker();
_ = new Tracker();
Console.WriteLine(Tracker.Instances);
"#
        ),
        &["1"]
    );
}

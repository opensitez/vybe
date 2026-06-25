//! Multicast delegates: `+=` combines handlers; `-=` removes by identity.
use super::helpers::run_csharp;

#[test]
fn multicast_delegate_invokes_handlers_in_subscription_order() {
    assert_eq!(
        run_csharp(
            r#"
using System;
void A() { Console.WriteLine("A"); }
void B() { Console.WriteLine("B"); }
Action chain = A;
chain += B;
chain();
"#
        ),
        &["A", "B"]
    );
}

#[test]
fn removing_one_handler_leaves_remaining_subscribers_active() {
    assert_eq!(
        run_csharp(
            r#"
using System;
void A() { Console.WriteLine("A"); }
void B() { Console.WriteLine("B"); }
Action chain = A;
chain += B;
chain -= A;
chain();
"#
        ),
        &["B"]
    );
}

#[test]
fn subtracting_absent_handler_is_silent_no_op() {
    assert_eq!(
        run_csharp(
            r#"
using System;
void A() { Console.WriteLine("A"); }
void B() { Console.WriteLine("B"); }
Action chain = A;
chain -= B;
chain();
"#
        ),
        &["A"]
    );
}

#[test]
fn func_multicast_combines_return_values_only_from_last_invoked() {
    assert_eq!(
        run_csharp(
            r#"
using System;
Func<int> first = () => { Console.WriteLine("1"); return 1; };
Func<int> second = () => { Console.WriteLine("2"); return 2; };
Func<int> chain = first;
chain += second;
Console.WriteLine(chain());
"#
        ),
        &["1", "2", "2"]
    );
}

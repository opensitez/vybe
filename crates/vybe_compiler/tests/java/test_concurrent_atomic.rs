/// java.util.concurrent.atomic — distinct atomic variable operations.
use crate::helpers::run_main;

#[test]
fn atomic_integer_get_and_increment_returns_prior_value() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(5); System.out.println(n.getAndIncrement()); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn atomic_integer_increment_and_get_returns_updated_value() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(9); System.out.println(n.incrementAndGet());",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn atomic_integer_add_and_get_accumulates_delta() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(3); System.out.println(n.addAndGet(4));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn atomic_integer_compare_and_set_succeeds_when_expected_matches() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(2); boolean ok = n.compareAndSet(2, 8); System.out.println(ok); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["true", "8"]);
}

#[test]
fn atomic_integer_compare_and_set_fails_when_expected_differs() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(2); boolean ok = n.compareAndSet(3, 8); System.out.println(ok); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["false", "2"]);
}

#[test]
fn atomic_integer_get_and_set_replaces_value() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(1); System.out.println(n.getAndSet(5)); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn atomic_boolean_compare_and_set_flips_false_to_true() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicBoolean flag = new java.util.concurrent.atomic.AtomicBoolean(false); System.out.println(flag.compareAndSet(false, true)); System.out.println(flag.get());",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn atomic_reference_swap_updates_stored_object() {
    let out = run_main(
        r#"java.util.concurrent.atomic.AtomicReference<String> ref = new java.util.concurrent.atomic.AtomicReference<String>("old"); System.out.println(ref.getAndSet("new")); System.out.println(ref.get());"#,
    );
    assert_eq!(out, vec!["old", "new"]);
}

#[test]
fn atomic_long_add_and_get_for_large_counter() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicLong n = new java.util.concurrent.atomic.AtomicLong(1_000_000L); System.out.println(n.addAndGet(250_000L));",
    );
    assert_eq!(out, vec!["1250000"]);
}

#[test]
fn atomic_integer_lazy_set_eventually_visible() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(0); n.lazySet(3); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn atomic_integer_weak_compare_and_set_updates_on_match() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(4); boolean ok = n.weakCompareAndSet(4, 6); System.out.println(ok); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["true", "6"]);
}

#[test]
fn atomic_integer_get_and_decrement_steps_down() {
    let out = run_main(
        "java.util.concurrent.atomic.AtomicInteger n = new java.util.concurrent.atomic.AtomicInteger(2); System.out.println(n.getAndDecrement()); System.out.println(n.get());",
    );
    assert_eq!(out, vec!["2", "1"]);
}

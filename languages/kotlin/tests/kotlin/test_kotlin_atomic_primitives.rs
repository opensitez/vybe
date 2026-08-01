use crate::helpers::run_prints;

#[test]
fn test_atomic_integer_basic_ops() {
    let out = run_prints(
        r#"
        fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(5)
            println(counter.get())
            println(counter.incrementAndGet())
            println(counter.getAndAdd(3))
            println(counter.get())
        }
    "#,
    );
    assert_eq!(out, &["5", "6", "6", "9"]);
}

#[test]
fn test_atomic_integer_compare_and_set() {
    let out = run_prints(
        r#"
        fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(10)
            val first = counter.compareAndSet(10, 20)
            val second = counter.compareAndSet(10, 30)
            println(first)
            println(second)
            println(counter.get())
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "20"]);
}

#[test]
fn test_atomic_long_update_and_get() {
    let out = run_prints(
        r#"
        fun main() {
            val value = java.util.concurrent.atomic.AtomicLong(0L)
            value.addAndGet(15)
            println(value.incrementAndGet())
            println(value.getAndSet(100))
            println(value.get())
        }
    "#,
    );
    assert_eq!(out, &["16", "16", "100"]);
}

#[test]
fn test_atomic_boolean_flip() {
    let out = run_prints(
        r#"
        fun main() {
            val flag = java.util.concurrent.atomic.AtomicBoolean(false)
            println(flag.compareAndSet(false, true))
            println(flag.getAndSet(false))
            println(flag.get())
            println(flag.compareAndSet(true, false))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "false"]);
}

#[test]
fn test_atomic_reference_set_and_exchange() {
    let out = run_prints(
        r#"
        fun main() {
            val ref = java.util.concurrent.atomic.AtomicReference("a")
            println(ref.get())
            println(ref.getAndSet("b"))
            println(ref.compareAndSet("b", "c"))
            println(ref.get())
        }
    "#,
    );
    assert_eq!(out, &["a", "a", "true", "c"]);
}

#[test]
fn test_atomic_reference_compare_and_set_failure() {
    let out = run_prints(
        r#"
        fun main() {
            val ref = java.util.concurrent.atomic.AtomicReference(5)
            val ok = ref.compareAndSet(4, 7)
            println(ok)
            println(ref.get())
        }
    "#,
    );
    assert_eq!(out, &["false", "5"]);
}

#[test]
fn test_atomic_update_with_lambda() {
    let out = run_prints(
        r#"
        fun main() {
            val counter = java.util.concurrent.atomic.AtomicInteger(1)
            val updated = counter.updateAndGet { value -> value * 3 }
            println(updated)
            val finalValue = counter.accumulateAndGet(4) { left, right -> left + right }
            println(finalValue)
        }
    "#,
    );
    assert_eq!(out, &["3", "7"]);
}

#[test]
fn test_atomic_lazy_set_and_weak_compare() {
    let out = run_prints(
        r#"
        fun main() {
            val value = java.util.concurrent.atomic.AtomicInteger(0)
            value.lazySet(9)
            println(value.get())
            val ok = value.compareAndSet(9, 10)
            println(ok)
            println(value.get())
        }
    "#,
    );
    assert_eq!(out, &["9", "true", "10"]);
}

#[test]
fn test_atomic_marking_sequence() {
    let out = run_prints(
        r#"
        import java.util.concurrent.atomic.AtomicInteger

        fun main() {
            val a = AtomicInteger(1)
            var state = ""
            repeat(3) {
                state += a.getAndIncrement().toString() + ","
            }
            println(state)
            println(a.get())
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,", "4"]);
}

#[test]
fn test_atomic_reference_object_identity() {
    let out = run_prints(
        r#"
        fun main() {
            data class PairBox(val value: Int)
            val a = PairBox(1)
            val b = PairBox(1)
            val ref = java.util.concurrent.atomic.AtomicReference(a)
            println(ref.compareAndSet(a, b))
            println(ref.get().value)
        }
    "#,
    );
    assert_eq!(out, &["true", "1"]);
}

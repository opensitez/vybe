use crate::helpers::run_prints;

#[test]
fn test_lazy_default_initializes_once() {
    let out = run_prints(
        r#"
        fun main() {
            var initCount = 0
            val x: Int by lazy {
                initCount++
                5 + 3
            }
            println(initCount)
            println(x)
            println(x)
            println(initCount)
        }
    "#,
    );
    assert_eq!(out, &["0", "8", "8", "1"]);
}

#[test]
fn test_lazy_thread_safety_modes() {
    let out = run_prints(
        r#"
        import kotlin.LazyThreadSafetyMode
        fun main() {
            var count = 0
            val a by lazy(LazyThreadSafetyMode.NONE) {
                count += 1
                "a"
            }
            val b by lazy(LazyThreadSafetyMode.SYNCHRONIZED) {
                count += 10
                "b"
            }
            println(a)
            println(a)
            println(b)
            println(b)
            println(count)
        }
    "#,
    );
    assert_eq!(out, &["a", "a", "b", "b", "11"]);
}

#[test]
fn test_lazy_returns_computation_result_once_even_with_side_effects() {
    let out = run_prints(
        r#"
        var sideEffect = 0
        fun load(): Int {
            sideEffect += 2
            return sideEffect
        }

        fun main() {
            val value by lazy { load() }
            println(value)
            println(value)
            println(sideEffect)
        }
    "#,
    );
    assert_eq!(out, &["2", "2", "2"]);
}

#[test]
fn test_lazy_property_as_declaration_output() {
    let out = run_prints(
        r#"
        class Holder {
            private val raw: Int = 3
            val doubled by lazy { raw * 2 }
        }

        fun main() {
            val h = Holder()
            println(h.doubled)
            println(h.doubled)
        }
    "#,
    );
    assert_eq!(out, &["6", "6"]);
}

#[test]
fn test_delegates_observable_records_changes() {
    let out = run_prints(
        r#"
        import kotlin.properties.Delegates

        fun main() {
            var events = ""
            var value by Delegates.observable(1) { _, old, new ->
                events += old.toString() + ":" + new.toString() + ";"
            }
            value = 2
            value = 5
            println(events)
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["1:2;2:5;", "5"]);
}

#[test]
fn test_delegates_vetoable_rejects_invalid_value() {
    let out = run_prints(
        r#"
        import kotlin.properties.Delegates

        fun main() {
            var value by Delegates.vetoable(1) { _, _, newValue ->
                newValue >= 0
            }
            value = 3
            println(value)
            value = -10
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_delegates_not_null_requires_assignment_before_read() {
    let out = run_prints(
        r#"
        import kotlin.properties.Delegates

        fun main() {
            class Holder {
                var name: String by Delegates.notNull()
            }

            val holder = Holder()
            try {
                holder.name.length
                println("ready")
            } catch (e: IllegalStateException) {
                println(e::class.simpleName)
            }
            holder.name = "ok"
            println(holder.name)
        }
    "#,
    );
    assert_eq!(out, &["IllegalStateException", "ok"]);
}

#[test]
fn test_lazy_uses_value_once_per_instance() {
    let out = run_prints(
        r#"
        fun main() {
            var count = 0
            class Holder {
                val value by lazy { count++ ; "x" }
            }
            val first = Holder()
            val second = Holder()
            println(first.value)
            println(first.value)
            println(second.value)
            println(count)
        }
    "#,
    );
    assert_eq!(out, &["x", "x", "x", "2"]);
}

#[test]
fn test_lazy_after_mutation_from_other_property() {
    let out = run_prints(
        r#"
        class Holder {
            var seed = 1
            val value by lazy { seed * 10 }
        }

        fun main() {
            val h = Holder()
            h.seed = 3
            println(h.value)
            h.seed = 9
            println(h.value)
        }
    "#,
    );
    assert_eq!(out, &["30", "30"]);
}

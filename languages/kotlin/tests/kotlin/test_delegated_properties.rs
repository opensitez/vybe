kotlin_run_test!(
    test_lazy_delegate_runs_once,
    r#"
        class Counter {
            var hits = 0
            val value by lazy { hits += 1; 42 }
        }

        fun main() {
            val counter = Counter()
            println(counter.value)
            println(counter.value)
            println(counter.hits)
        }
    "#,
    &["42", "42", "1"]
);

kotlin_run_test!(
    test_lazy_custom_policy_none,
    r#"
        import kotlin.LazyThreadSafetyMode

        class Holder {
            var invoked = 0
            val value by lazy(LazyThreadSafetyMode.NONE) {
                invoked += 1
                "done"
            }
        }

        fun main() {
            val h = Holder()
            println(h.value)
            println(h.value)
            println(h.invoked)
        }
    "#,
    &["done", "done", "1"]
);

kotlin_run_test!(
    test_observable_tracks_changes,
    r#"
        import kotlin.properties.Delegates

        fun main() {
            val events = mutableListOf<String>()
            var value by Delegates.observable("init") { _, old, new ->
                events.add(old + ">" + new)
            }
            value = "a"
            value = "b"
            println(value)
            println(events.joinToString(","))
        }
    "#,
    &["b", "init>a,a>b"]
);

kotlin_run_test!(
    test_vetoable_rejects_negative_value,
    r#"
        import kotlin.properties.Delegates

        fun main() {
            var score by Delegates.vetoable(1) { _, old, new ->
                new >= 0
            }
            score = -3
            val first = score
            score = 7
            val second = score
            println(first)
            println(second)
        }
    "#,
    &["1", "7"]
);

kotlin_run_test!(
    test_not_null_delegate_requires_initialization,
    r#"
        import kotlin.properties.Delegates

        class Box {
            var value: Int by Delegates.notNull()
        }

        fun main() {
            val box = Box()
            val out = try {
                box.value
                "ok"
            } catch (e: IllegalStateException) {
                "not-set"
            }
            println(out)
            box.value = 11
            println(box.value)
        }
    "#,
    &["not-set", "11"]
);

kotlin_run_test!(
    test_observable_with_no_change,
    r#"
        import kotlin.properties.Delegates

        fun main() {
            val events = mutableListOf<String>()
            var value by Delegates.observable(10) { _, old, new ->
                events.add("${'$'}old/${'$'}new")
            }
            value = 12
            value = 12
            println(events.size)
            println(events[0])
            println(events[1])
        }
    "#,
    &["2", "10/12", "12/12"]
);

kotlin_run_test!(
    test_vetoable_allows_true_transition,
    r#"
        import kotlin.properties.Delegates

        var total by Delegates.vetoable(0) { _, _, new -> new >= 0 }

        fun main() {
            total = 5
            total = 6
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_lazy_reuses_cached_value_after_external_mutation,
    r#"
        class Cache {
            private var calls = 0
            val value by lazy {
                calls += 1
                calls * 3
            }
            fun currentCalls(): Int = calls
        }

        fun main() {
            val c = Cache()
            println(c.value)
            println(c.value)
            println(c.currentCalls())
        }
    "#,
    &["3", "3", "1"]
);

kotlin_run_test!(
    test_observable_can_record_multiple_types,
    r#"
        import kotlin.properties.Delegates

        class Tracker {
            var events = 0
            var label by Delegates.observable("x") { _, old, new ->
                if (old != new) events += 1
            }
        }

        fun main() {
            val t = Tracker()
            t.label = "a"
            t.label = "a"
            t.label = "b"
            println(t.label)
            println(t.events)
        }
    "#,
    &["b", "2"]
);

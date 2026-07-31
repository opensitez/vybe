kotlin_run_test!(
    test_use_closes_marker_on_success,
    r#"
        class Marker : AutoCloseable {
            var closed = false
            override fun close() {
                closed = true
            }
        }

        fun main() {
            val marker = Marker()
            marker.use {
                println("open")
            }
            println(marker.closed)
        }
    "#,
    &["open", "true"]
);

kotlin_run_test!(
    test_try_finally_always_runs_close,
    r#"
        class Marker {
            var closed = false
            fun close() { closed = true }
        }

        fun main() {
            val marker = Marker()
            try {
                println("inside")
            } finally {
                marker.close()
            }
            println(marker.closed)
        }
    "#,
    &["inside", "true"]
);

kotlin_run_test!(
    test_nested_uses_close_outer_and_inner,
    r#"
        class Token(val name: String) : AutoCloseable {
            var calls: Int = 0
            override fun close() {
                calls += 1
            }
        }

        fun main() {
            val a = Token("a")
            val b = Token("b")
            a.use {
                b.use {
                    println(a.calls + b.calls)
                }
            }
            println(a.calls)
            println(b.calls)
        }
    "#,
    &["0", "1", "1"]
);

kotlin_run_test!(
    test_try_catch_still_runs_close,
    r#"
        class Token : AutoCloseable {
            var closed = false
            override fun close() {
                closed = true
            }
        }

        fun main() {
            val token = Token()
            try {
                throw IllegalStateException("x")
            } catch (_: IllegalStateException) {
                println("err")
            } finally {
                token.close()
            }
            println(token.closed)
        }
    "#,
    &["err", "true"]
);

kotlin_run_test!(
    test_manual_with_resource_guard_and_result,
    r#"
        class Counter : AutoCloseable {
            var total = 0
            override fun close() {
                total += 10
            }
        }

        fun main() {
            val c = Counter()
            c.use {
                it.total = 5
                println(it.total)
            }
            println(c.total)
        }
    "#,
    &["5", "15"]
);

kotlin_run_test!(
    test_resource_on_collection_mapping,
    r#"
        class LogClose : AutoCloseable {
            var tag = ""
            override fun close() { tag = "closed" }
        }

        fun main() {
            val logs = listOf("a", "b")
            val out = logs.joinToString(",") { value ->
                val resource = LogClose()
                resource.use {
                    println(value)
                    value
                }
            }
            println(out)
        }
    "#,
    &["a", "b", "a,b"]
);

kotlin_run_test!(
    test_resource_acquired_before_body_and_closed_after,
    r#"
        class Holder : AutoCloseable {
            var active = false
            override fun close() { active = false }
        }

        fun main() {
            var h: Holder? = null
            Holder().use {
                it.active = true
                h = it
                println(it.active)
            }
            println(h?.active ?: false)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_resource_in_early_returned_block,
    r#"
        class Slot : AutoCloseable {
            var closed = false
            override fun close() { closed = true }
        }

        fun compute(v: Int): String {
            val slot = Slot()
            slot.use {
                if (v == 0) return "zero"
            }
            return "done"
        }

        fun main() {
            println(compute(0))
        }
    "#,
    &["zero"]
);

kotlin_run_test!(
    test_resource_multiple_scopes_are_independent,
    r#"
        class Flag : AutoCloseable {
            var closeCount = 0
            override fun close() { closeCount += 1 }
        }

        fun main() {
            val first = Flag()
            val second = Flag()
            first.use { }
            second.use { }
            println(first.closeCount)
            println(second.closeCount)
        }
    "#,
    &["1", "1"]
);

kotlin_run_test!(
    test_try_finally_overrides_return,
    r#"
        fun main() {
            var out = "init"
            try {
                out = "inside"
                return@main
            } finally {
                out = "final"
                println(out)
            }
        }
    "#,
    &["final"]
);

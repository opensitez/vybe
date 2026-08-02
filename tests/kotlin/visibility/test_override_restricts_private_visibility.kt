// vybe-test: kotlin/visibility/test_override_restricts_private_visibility
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            private open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
        }


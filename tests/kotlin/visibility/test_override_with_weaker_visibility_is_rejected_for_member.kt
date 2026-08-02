// vybe-test: kotlin/visibility/test_override_with_weaker_visibility_is_rejected_for_member
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            protected open fun label(): String = "base"
        }

        class Child : Base() {
            private override fun label(): String = "child"
        }


// vybe-test: kotlin/visibility/test_private_member_stays_within_class_hierarchy
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            private val hidden = 4
        }

        class Child : Base() {
            fun read(): Int = hidden
        }


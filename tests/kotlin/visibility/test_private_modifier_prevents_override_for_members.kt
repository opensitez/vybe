// vybe-test: kotlin/visibility/test_private_modifier_prevents_override_for_members
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            private fun hidden() = "x"
        }

        class Child : Base() {
            fun call(base: Base): String {
                return base.hidden()
            }
        }


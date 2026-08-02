// vybe-test: kotlin/visibility/test_private_constructor_hides_instantiation_outside_factory
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Locked private constructor()

        fun main() {
            Locked()
        }


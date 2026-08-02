// vybe-test: kotlin/visibility/test_private_setter_is_not_accessible_for_extension_receiver
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Counter {
            var value: Int = 0
                private set
        }

        fun Counter.bump() {
            this.value = this.value + 1
        }

        fun main() {
            Counter().bump()
        }


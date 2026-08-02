// vybe-test: kotlin/visibility/test_private_setter_restricts_external_assignment
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Counter {
            var value: Int = 0
                private set
        }

        fun main() {
            val counter = Counter()
            counter.value = 4
            println(counter.value)
        }


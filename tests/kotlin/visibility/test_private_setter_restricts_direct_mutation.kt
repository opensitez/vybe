// vybe-test: kotlin/visibility/test_private_setter_restricts_direct_mutation
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Count {
            var value: Int = 0
                private set
        }

        fun main() {
            val count = Count()
            count.value = 4
            println(count.value)
        }


// vybe-test: kotlin/scoping_functions/test_take_unless_predicate_not_called_for_null_receiver
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun main() {
            val value: Int? = null
            val result = value?.takeUnless {
                println("should-not-see-this")
                false
            }
            println(result == null)
        }


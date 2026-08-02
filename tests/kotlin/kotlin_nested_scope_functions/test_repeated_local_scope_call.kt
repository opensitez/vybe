// vybe-test: kotlin/kotlin_nested_scope_functions/test_repeated_local_scope_call
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun main() {
            val values = mutableListOf<Int>()
            run {
                for (i in 1..3) values.add(i)
            }
            println(values.joinToString(""))
        }


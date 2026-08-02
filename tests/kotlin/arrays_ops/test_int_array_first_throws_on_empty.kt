// vybe-test: kotlin/arrays_ops/test_int_array_first_throws_on_empty
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val none = intArrayOf()
            try {
                println(none.first())
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }


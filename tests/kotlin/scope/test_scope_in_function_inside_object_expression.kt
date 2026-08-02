// vybe-test: kotlin/scope/test_scope_in_function_inside_object_expression
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            val maker = object : Any() {
                fun add(value: Int): Int {
                    fun inner(): Int {
                        return value + 1
                    }
                    return inner()
                }
            }
            total += maker.add(2)
            __check((total).toString(), "3")
        }

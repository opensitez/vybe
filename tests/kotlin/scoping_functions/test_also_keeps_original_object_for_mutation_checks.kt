// vybe-test: kotlin/scoping_functions/test_also_keeps_original_object_for_mutation_checks
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder(var total: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = Holder(5)
            val observed = original.also {
                it.total += 10
            }
            __check((original.total).toString(), "15")
            __check((original === observed).toString(), "true")
        }

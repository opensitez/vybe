// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_array_mutation_in_lambda
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = IntArray(3)
            val mutate = { idx: Int, value: Int -> list[idx] = value }
            mutate(0, 5)
            mutate(1, 6)
            __check((list.joinToString(",")).toString(), "5,6,0")
        }

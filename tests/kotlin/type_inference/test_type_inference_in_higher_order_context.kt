// vybe-test: kotlin/type_inference/test_type_inference_in_higher_order_context
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun build(fn: (Int) -> Int): Int = fn(4)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = build { it + 1 }
            __check((x).toString(), "5")
        }

// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_and_return_factory
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun makeAdder(base: Int): (Int) -> Int {
            return { x -> x + base }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val add5 = makeAdder(5)
            __check((add5(10)).toString(), "15")
        }

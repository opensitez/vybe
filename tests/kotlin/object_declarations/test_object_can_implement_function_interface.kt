// vybe-test: kotlin/object_declarations/test_object_can_implement_function_interface
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Incrementer : (Int) -> Int {
            override fun invoke(value: Int): Int = value + 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Incrementer(2)).toString(), "3")
            __check((Incrementer.invoke(4)).toString(), "5")
        }

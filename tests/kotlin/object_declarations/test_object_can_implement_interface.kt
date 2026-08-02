// vybe-test: kotlin/object_declarations/test_object_can_implement_interface
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Handler {
            fun call(value: Int): Int
        }

        object PlusOne : Handler {
            override fun call(value: Int): Int = value + 1
        }

        fun apply(handler: Handler, value: Int): Int = handler.call(value)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(PlusOne, 4)).toString(), "5")
        }

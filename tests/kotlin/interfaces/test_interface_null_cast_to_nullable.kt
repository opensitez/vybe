// vybe-test: kotlin/interfaces/test_interface_null_cast_to_nullable
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Reader { fun read(): Int }

        class NumberReader : Reader {
            override fun read(): Int = 7
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Any? = null
            val value = source as? Reader
            __check((value == null).toString(), "true")
            __check(((NumberReader() as Reader).read()).toString(), "7")
        }

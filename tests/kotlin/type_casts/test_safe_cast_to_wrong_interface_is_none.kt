// vybe-test: kotlin/type_casts/test_safe_cast_to_wrong_interface_is_none
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

interface Reader { fun read(): String }
        class FileReader : Reader {
            override fun read(): String = "ok"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 99
            val reader = value as? Reader
            __check((reader == null).toString(), "true")
        }

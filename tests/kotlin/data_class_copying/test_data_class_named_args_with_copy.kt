// vybe-test: kotlin/data_class_copying/test_data_class_named_args_with_copy
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Config(val host: String, val port: Int, val secure: Boolean)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Config("localhost", 80, false)
            val secure = base.copy(port = 443, secure = true)
            __check((secure.host).toString(), "localhost")
            __check((secure.port).toString(), "443")
            __check((secure.secure).toString(), "true")
        }

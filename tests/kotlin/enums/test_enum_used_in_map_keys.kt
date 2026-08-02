// vybe-test: kotlin/enums/test_enum_used_in_map_keys
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Option { READ, WRITE, EXECUTE }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val permissions = mapOf(
                Option.READ to "read-only",
                Option.WRITE to "read-write",
                Option.EXECUTE to "admin"
            )
            __check((permissions[Option.READ]).toString(), "read-only")
            __check((permissions[Option.EXECUTE]).toString(), "admin")
        }

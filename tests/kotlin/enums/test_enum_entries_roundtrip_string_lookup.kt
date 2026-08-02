// vybe-test: kotlin/enums/test_enum_entries_roundtrip_string_lookup
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Mode { READ, WRITE, EXECUTE }

        fun describe(value: String): String {
            return try {
                val mode = Mode.valueOf(value)
                "ok:" + mode.name
            } catch (e: Exception) {
                "bad"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("WRITE")).toString(), "ok:WRITE")
            __check((describe("bad")).toString(), "bad")
        }

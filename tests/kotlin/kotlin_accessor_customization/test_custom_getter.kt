// vybe-test: kotlin/kotlin_accessor_customization/test_custom_getter
// origin: languages/kotlin/tests/kotlin/test_kotlin_accessor_customization.rs

class Score {
            private var raw = 0

            var value: Int
                get() = raw * 2
                set(v) { raw = if (v < 0) 0 else v }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = Score()
            s.value = 3
            __check((s.value).toString(), "6")
            s.value = -4
            __check((s.value).toString(), "0")
        }

// vybe-test: kotlin/kotlin_visibility_keywords/test_local_visibility_with_private_setter
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_keywords.rs

class Data {
            var value: Int = 0
                private set

            fun assign(next: Int) {
                value = next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val d = Data()
            d.assign(3)
            __check((d.value).toString(), "3")
        }

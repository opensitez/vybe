// vybe-test: kotlin/visibility/test_public_setter_allows_external_mutation
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Count {
            var value: Int = 0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val count = Count()
            count.value = 4
            __check((count.value).toString(), "4")
        }

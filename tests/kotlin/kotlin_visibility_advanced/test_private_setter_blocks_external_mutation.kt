// vybe-test: kotlin/kotlin_visibility_advanced/test_private_setter_blocks_external_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

class Counter {
            var value: Int = 0
                private set

            fun bump() { value += 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            counter.bump()
            counter.bump()
            __check((counter.value).toString(), "2")
        }

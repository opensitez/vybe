// vybe-test: kotlin/kotlin_property_accessors_advanced/test_custom_indexing_like_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Bag {
            private val data = listOf(1, 3, 5)
            operator fun get(index: Int): Int = data[index]
            val head get() = data.first()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Bag()
            __check((b[2]).toString(), "5")
            __check((b.head).toString(), "1")
        }

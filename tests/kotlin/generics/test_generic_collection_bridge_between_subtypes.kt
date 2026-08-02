// vybe-test: kotlin/generics/test_generic_collection_bridge_between_subtypes
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> mergeInto(dest: MutableList<T>, first: T, second: T) {
            dest.add(first)
            dest.add(second)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = mutableListOf<Any>()
            mergeInto(data, 1, "x")
            __check((data.size).toString(), "2")
            __check((data[0]).toString(), "1")
            __check((data[1]).toString(), "x")
        }

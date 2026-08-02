// vybe-test: kotlin/kotlin_set_apis/test_set_to_mutable_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val immutable = setOf(1, 2)
            val mutable = immutable.toMutableSet()
            mutable.add(3)
            __check((immutable.size).toString(), "2")
            __check((mutable.size).toString(), "3")
            __check((mutable.contains(3)).toString(), "true")
        }

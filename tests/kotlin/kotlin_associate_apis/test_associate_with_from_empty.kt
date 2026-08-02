// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_from_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf<String>().associateWith { it.length }
            __check((map.isEmpty()).toString(), "true")
            __check((map.size).toString(), "0")
        }

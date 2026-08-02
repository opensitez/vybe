// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_transform_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf("x", "yy", "zzz")
            val map = list.associateBy({ it }, { it.length })
            __check((map["x"]).toString(), "1")
            __check((map["zzz"]).toString(), "3")
        }

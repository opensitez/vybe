// vybe-test: kotlin/exceptions/test_exception_finally_does_not_modify_returned_value_binding
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun compute(): Int {
    var result = 1
    try {
        __check(("try").toString(), "try")
        result = 5
        return result
    } finally {
        __check(("finally").toString(), "finally")
        result = 9
    }
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    __check((compute()).toString(), "5")
}

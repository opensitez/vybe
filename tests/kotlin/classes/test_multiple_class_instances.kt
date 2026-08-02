// vybe-test: kotlin/classes/test_multiple_class_instances
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Account(val id: String, var balance: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a1 = Account("A", 100)
            val a2 = Account("B", 200)
            a1.balance += 50
            __check((a1.balance).toString(), "150")
            __check((a2.balance).toString(), "200")
        }

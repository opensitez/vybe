// vybe-test: kotlin/classes/test_private_like_state_mutation
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Bank { var balance: Int = 100
fun deposit(v: Int) { balance += v }
fun withdraw(v: Int) { balance -= v }
fun total(): Int = balance }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val b = Bank()
b.deposit(40)
b.withdraw(10)
__check((b.total()).toString(), "130") }

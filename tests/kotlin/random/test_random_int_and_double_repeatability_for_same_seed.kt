// vybe-test: kotlin/random/test_random_int_and_double_repeatability_for_same_seed
// origin: languages/kotlin/tests/kotlin/test_random.rs

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = kotlin.random.Random(113)
            val b = kotlin.random.Random(113)
            val aFirst = a.nextInt(1000)
            val aSecond = a.nextInt(1000)
            val bFirst = b.nextInt(1000)
            val bSecond = b.nextInt(1000)
            val aDouble = a.nextDouble()
            val bDouble = b.nextDouble()
            __p((aFirst == bFirst).toString())
            __p((aSecond == bSecond).toString())
            __p((aDouble == bDouble).toString())
        
__check("true\ntrue\ntrue")
}

// vybe-test: kotlin/this_super/test_this_outer_in_class_nesting
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class Level1 {
        val x = "one"
        class Level2(val parent: Level1) { fun read() = parent.x }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val p = Level1()
__check((Level1.Level2(p).read()).toString(), "one") }

// vybe-test: kotlin/bitwise_operations/test_flags_with_union_and_intersection
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val canRead = 0b001
            val canWrite = 0b010
            val canExecute = 0b100
            val perms = canRead or canWrite
            __check((perms and canRead).toString(), "1")
            __check((perms and canExecute).toString(), "0")
            val withExec = perms or canExecute
            val withoutRead = withExec and canRead.inv()
            __check((withExec).toString(), "7")
            __check((withoutRead).toString(), "-8")
            __check((withoutRead and canWrite).toString(), "2")
            __check((withoutRead and canExecute).toString(), "4")
        }

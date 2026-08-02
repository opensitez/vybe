// vybe-test: kotlin/this_super/test_this_inside_try_block
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K { fun id() = this.toString() }
fun main() { try { println(K().id().isNotEmpty()) } catch (e: Exception) { println("err") } }


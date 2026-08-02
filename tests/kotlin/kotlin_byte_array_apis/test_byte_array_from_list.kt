// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_from_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun main() {
            val list = listOf<Byte>(1, 2, 3)
            val a = ByteArray(list.size) { list[it] }
            var out = ""
            for (x in a) { out = out + x.toString() }
            println(out)
        }


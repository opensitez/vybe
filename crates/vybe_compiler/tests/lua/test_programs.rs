//! Small programs combining control flow, functions, and data (integration gaps).

lua_print! {
    sum_first_n_integers => {
        "local n = 5\nlocal sum = 0\nlocal i = 1\nwhile i <= n do\n  sum = sum + i\n  i = i + 1\nend\nprint(sum)\n",
        "15"
    },
    fizzbuzz_one_line_for_fifteen => {
        "local n = 15\nif n % 15 == 0 then print(\"fizzbuzz\")\nelseif n % 3 == 0 then print(\"fizz\")\nelseif n % 5 == 0 then print(\"buzz\")\nelse print(n)\nend\n",
        "fizzbuzz"
    },
    collatz_step_count_to_one => {
        "local n = 10\nlocal steps = 0\nwhile n ~= 1 do\n  if n % 2 == 0 then n = n // 2 else n = 3 * n + 1 end\n  steps = steps + 1\nend\nprint(steps)\n",
        "6"
    },
    euclidean_gcd_via_modulo => {
        "local a, b = 48, 18\nwhile b ~= 0 do\n  a, b = b, a % b\nend\nprint(a)\n",
        "6"
    },
    insertion_sort_on_small_array => {
        "local t = {3, 1, 4, 2}\nfor i = 2, #t do\n  local key = t[i]\n  local j = i - 1\n  while j > 0 and t[j] > key do\n    t[j + 1] = t[j]\n    j = j - 1\n  end\n  t[j + 1] = key\nend\nprint(table.concat(t, \",\"))\n",
        "1,2,3,4"
    },
    balanced_parentheses_checker => {
        "local s = \"(()())\"\nlocal depth = 0\nlocal ok = true\nfor i = 1, #s do\n  local c = string.sub(s, i, i)\n  if c == \"(\" then depth = depth + 1 elseif c == \")\" then depth = depth - 1 end\n  if depth < 0 then ok = false break end\nend\nprint(ok and depth == 0)\n",
        "true"
    },
    run_length_encoding_counts => {
        "local s = \"aaabbc\"\nlocal out = \"\"\nlocal i = 1\nwhile i <= #s do\n  local c = string.sub(s, i, i)\n  local n = 1\n  while string.sub(s, i + n, i + n) == c do n = n + 1 end\n  out = out .. c .. n\n  i = i + n\nend\nprint(out)\n",
        "a3b2c1"
    },
    tower_of_hanoi_move_count => {
        "local function hanoi(n)\n  if n == 1 then return 1 end\n  return 2 * hanoi(n - 1) + 1\nend\nprint(hanoi(4))\n",
        "15"
    },
    binary_search_finds_index => {
        "local t = {1, 3, 5, 7, 9}\nlocal target = 7\nlocal lo, hi = 1, #t\nlocal found = 0\nwhile lo <= hi do\n  local mid = (lo + hi) // 2\n  if t[mid] == target then found = mid break\n  elseif t[mid] < target then lo = mid + 1\n  else hi = mid - 1 end\nend\nprint(found)\n",
        "4"
    },
    dutch_flag_partition_zeros_and_ones => {
        "local t = {1, 0, 1, 0, 1}\nlocal low, high = 1, #t\nlocal i = 1\nwhile i <= high do\n  if t[i] == 0 then\n    t[low], t[i] = t[i], t[low]\n    low = low + 1\n    i = i + 1\n  else\n    high = high - 1\n    t[i], t[high] = t[high], t[i]\n  end\nend\nprint(table.concat(t, \",\"))\n",
        "0,0,1,1,1"
    },
    sieve_of_eratosthenes_counts_primes_up_to_twenty => {
        "local n = 20\nlocal is_prime = {}\nfor i = 2, n do is_prime[i] = true end\nfor p = 2, math.floor(math.sqrt(n)) do\n  if is_prime[p] then\n    for m = p * p, n, p do is_prime[m] = false end\n  end\nend\nlocal count = 0\nfor i = 2, n do if is_prime[i] then count = count + 1 end end\nprint(count)\n",
        "8"
    },
    reverse_words_in_string => {
        "local s = \"one two three\"\nlocal rev = \"\"\nfor word in string.gmatch(s, \"%S+\") do\n  rev = word .. (rev == \"\" and \"\" or \" \") .. rev\nend\nprint(rev)\n",
        "three two one"
    },
    depth_first_traversal_sum_tree => {
        "local tree = {v=1, left={v=2}, right={v=3, right={v=4}}}\nlocal function sum(node)\n  if node == nil then return 0 end\n  return node.v + sum(node.left) + sum(node.right)\nend\nprint(sum(tree))\n",
        "10"
    },
    count_vowels_in_string => {
        "local s = \"hello\"\nlocal n = 0\nfor i = 1, #s do\n  local c = string.sub(s, i, i)\n  if c == \"a\" or c == \"e\" or c == \"i\" or c == \"o\" or c == \"u\" then n = n + 1 end\nend\nprint(n)\n",
        "2"
    },
    find_minimum_in_array => {
        "local t = {4, 1, 9, 2}\nlocal min = t[1]\nfor i = 2, #t do if t[i] < min then min = t[i] end end\nprint(min)\n",
        "1"
    },
    find_maximum_in_array => {
        "local t = {4, 1, 9, 2}\nlocal max = t[1]\nfor i = 2, #t do if t[i] > max then max = t[i] end end\nprint(max)\n",
        "9"
    },
    iterative_factorial => {
        "local n = 5\nlocal acc = 1\nfor i = 2, n do acc = acc * i end\nprint(acc)\n",
        "120"
    },
    iterative_fibonacci_nth => {
        "local n = 6\nlocal a, b = 0, 1\nfor i = 1, n do a, b = b, a + b end\nprint(a)\n",
        "8"
    },
    palindrome_string_check => {
        "local s = \"aba\"\nlocal ok = true\nfor i = 1, #s do\n  if string.sub(s, i, i) ~= string.sub(s, #s - i + 1, #s - i + 1) then ok = false break end\nend\nprint(ok)\n",
        "true"
    },
    absolute_difference_without_math_abs => {
        "local a, b = 3, 10\nlocal d = a - b\nif d < 0 then d = -d end\nprint(d)\n",
        "7"
    },
    clamp_value_between_bounds => {
        "local v = 15\nlocal lo, hi = 0, 10\nif v < lo then v = lo elseif v > hi then v = hi end\nprint(v)\n",
        "10"
    },
    average_of_three_numbers => {
        "local a, b, c = 2, 4, 6\nprint((a + b + c) / 3)\n",
        "4"
    },
    double_each_array_element => {
        "local t = {1, 2, 3}\nfor i = 1, #t do t[i] = t[i] * 2 end\nprint(table.concat(t, \",\"))\n",
        "2,4,6"
    },
    count_value_occurrences_in_array => {
        "local t = {1, 2, 1, 3, 1}\nlocal target = 1\nlocal n = 0\nfor i = 1, #t do if t[i] == target then n = n + 1 end end\nprint(n)\n",
        "3"
    },
    comma_split_into_array_manual => {
        "local s = \"a,b,c\"\nlocal t = {}\nfor part in string.gmatch(s, \"[^,]+\") do table.insert(t, part) end\nprint(table.concat(t, \"/\"))\n",
        "a/b/c"
    },
    starts_with_prefix_check => {
        "local s = \"lua.lang\"\nprint(string.sub(s, 1, 3) == \"lua\")\n",
        "true"
    },
    ends_with_suffix_check => {
        "local s = \"file.lua\"\nprint(string.sub(s, -4) == \".lua\")\n",
        "true"
    },
    replace_spaces_with_dashes => {
        "print(string.gsub(\"a b c\", \" \", \"-\"))\n",
        "a-b-c"
    },
    count_words_in_sentence => {
        "local n = 0\nfor _ in string.gmatch(\"one two three\", \"%S+\") do n = n + 1 end\nprint(n)\n",
        "3"
    },
    sum_of_squares => {
        "local t = {1, 2, 3}\nlocal s = 0\nfor i = 1, #t do s = s + t[i] * t[i] end\nprint(s)\n",
        "14"
    },
    product_of_list => {
        "local t = {2, 3, 4}\nlocal p = 1\nfor i = 1, #t do p = p * t[i] end\nprint(p)\n",
        "24"
    },
    linear_search_found_index => {
        "local t = {5, 9, 3}\nlocal target = 9\nlocal idx = 0\nfor i = 1, #t do if t[i] == target then idx = i break end end\nprint(idx)\n",
        "2"
    },
    reverse_array_in_place => {
        "local t = {1, 2, 3, 4}\nlocal i, j = 1, #t\nwhile i < j do t[i], t[j] = t[j], t[i] i = i + 1 j = j - 1 end\nprint(table.concat(t, \",\"))\n",
        "4,3,2,1"
    },
    count_zeros_in_list => {
        "local t = {0, 1, 0, 2, 0}\nlocal n = 0\nfor i = 1, #t do if t[i] == 0 then n = n + 1 end end\nprint(n)\n",
        "3"
    },
    iterative_integer_power => {
        "local base, exp = 2, 5\nlocal r = 1\nfor i = 1, exp do r = r * base end\nprint(r)\n",
        "32"
    },
    sign_function_with_if => {
        "local function sign(n)\n  if n > 0 then return 1 elseif n < 0 then return -1 else return 0 end\nend\nprint(sign(-2))\n",
        "-1"
    },
    max_of_two_without_math_lib => {
        "local a, b = 3, 7\nprint(a > b and a or b)\n",
        "7"
    },
    min_of_two_without_math_lib => {
        "local a, b = 3, 7\nprint(a < b and a or b)\n",
        "3"
    },
    parity_even_check => {
        "local n = 4\nprint(n % 2 == 0)\n",
        "true"
    },
    triangle_number_formula => {
        "local n = 4\nprint(n * (n + 1) // 2)\n",
        "10"
    },
    digit_sum_of_integer => {
        "local n = 123\nlocal s = 0\nwhile n > 0 do\n  s = s + (n % 10)\n  n = n // 10\nend\nprint(s)\n",
        "6"
    },
    greet_with_name_parameter => {
        "local function greet(name) return \"hi \" .. name end\nprint(greet(\"ada\"))\n",
        "hi ada"
    },
    filter_positive_into_new_list => {
        "local src = {-1, 2, -3, 4}\nlocal out = {}\nfor i = 1, #src do if src[i] > 0 then out[#out + 1] = src[i] end end\nprint(table.concat(out, \",\"))\n",
        "2,4"
    },
    accumulate_running_total_in_loop => {
        "local t = {1, 2, 3}\nlocal sum = 0\nfor i = 1, #t do sum = sum + t[i] end\nprint(sum)\n",
        "6"
    },
    copy_array_elements_to_new_table => {
        "local src = {1, 2}\nlocal dst = {}\nfor i = 1, #src do dst[i] = src[i] end\nprint(dst[2])\n",
        "2"
    },
    stack_push_pop_using_table => {
        "local st = {}\ntable.insert(st, 1)\ntable.insert(st, 2)\nprint(table.remove(st))\n",
        "2"
    },
    map_strings_to_upper_list => {
        "local t = {\"a\", \"b\"}\nfor i = 1, #t do t[i] = string.upper(t[i]) end\nprint(table.concat(t, \"\"))\n",
        "AB"
    },
    first_matching_predicate_index => {
        "local t = {2, 4, 6, 7}\nlocal idx = 0\nfor i = 1, #t do if t[i] % 2 == 1 then idx = i break end end\nprint(idx)\n",
        "4"
    },
    median_of_three_values => {
        "local a, b, c = 3, 1, 2\nif a > b then a, b = b, a end\nif b > c then b, c = c, b end\nif a > b then a, b = b, a end\nprint(b)\n",
        "2"
    },
    swap_first_and_last_elements => {
        "local t = {10, 20, 30}\nt[1], t[#t] = t[#t], t[1]\nprint(t[1] .. \",\" .. t[#t])\n",
        "30,10"
    },
    while_read_until_sentinel_pattern => {
        "local data = {1, 2, -1, 9}\nlocal sum = 0\nlocal i = 1\nwhile data[i] ~= -1 do sum = sum + data[i] i = i + 1 end\nprint(sum)\n",
        "3"
    },
    repeat_menu_until_quit_flag => {
        "local choice = 0\nlocal guard = 0\nrepeat\n  guard = guard + 1\n  choice = guard\nuntil choice >= 2\nprint(choice)\n",
        "2"
    },
    running_sum_with_numeric_for => {
        "local sum, n = 0, 5\nfor i = 1, n do sum = sum + i end\nprint(sum)\n",
        "15"
    },
    filter_even_numbers_into_new_list => {
        "local src = {1, 2, 3, 4, 5}\nlocal out = {}\nfor _, v in ipairs(src) do if v % 2 == 0 then table.insert(out, v) end end\nprint(#out)\n",
        "2"
    },
    map_double_values_with_ipairs => {
        "local t = {2, 3, 4}\nfor i, v in ipairs(t) do t[i] = v * 2 end\nprint(t[2])\n",
        "6"
    },
    reduce_product_of_list => {
        "local t = {2, 3, 4}\nlocal p = 1\nfor _, v in ipairs(t) do p = p * v end\nprint(p)\n",
        "24"
    },
    count_hash_entries_with_pairs => {
        "local cfg = {host = \"x\", port = 80, ssl = true}\nlocal n = 0\nfor _ in pairs(cfg) do n = n + 1 end\nprint(n)\n",
        "3"
    },
    closure_counter_in_loop_body => {
        "local fns = {}\nfor i = 1, 3 do\n  fns[i] = function() return i end\nend\nprint(fns[2]())\n",
        "2"
    },
    higher_order_apply_twice => {
        "local function apply_twice(f, x) return f(f(x)) end\nprint(apply_twice(function(n) return n + 1 end, 3))\n",
        "5"
    },
    compose_two_unary_functions => {
        "local function compose(f, g) return function(x) return f(g(x)) end end\nlocal inc = function(n) return n + 1 end\nlocal double = function(n) return n * 2 end\nprint(compose(inc, double)(3))\n",
        "7"
    },
    memoized_fibonacci_table => {
        "local cache = {[0] = 0, [1] = 1}\nlocal function fib(n)\n  if cache[n] then return cache[n] end\n  cache[n] = fib(n - 1) + fib(n - 2)\n  return cache[n]\nend\nprint(fib(8))\n",
        "21"
    },
    queue_enqueue_dequeue_fifo => {
        "local q = {}\nlocal head, tail = 1, 0\nlocal function enqueue(v) tail = tail + 1 q[tail] = v end\nlocal function dequeue() local v = q[head] head = head + 1 return v end\nenqueue(1) enqueue(2)\nprint(dequeue() + dequeue())\n",
        "3"
    },
    binary_search_in_sorted_array => {
        "local t = {1, 3, 5, 7, 9}\nlocal target, lo, hi = 7, 1, #t\nwhile lo <= hi do\n  local mid = math.floor((lo + hi) / 2)\n  if t[mid] == target then print(mid) return end\n  if t[mid] < target then lo = mid + 1 else hi = mid - 1 end\nend\nprint(0)\n",
        "4"
    },
    bubble_sort_small_array => {
        "local t = {3, 1, 2}\nfor i = 1, #t - 1 do\n  for j = 1, #t - i do\n    if t[j] > t[j + 1] then t[j], t[j + 1] = t[j + 1], t[j] end\n  end\nend\nprint(table.concat(t, \",\"))\n",
        "1,2,3"
    },
    string_builder_via_concat_loop => {
        "local parts = {\"a\", \"b\", \"c\"}\nlocal s = \"\"\nfor i = 1, #parts do s = s .. parts[i] end\nprint(s)\n",
        "abc"
    },
    truthy_zero_in_conditional_branch => {
        "local n = 0\nif n then print(\"truthy\") else print(\"falsy\") end\n",
        "truthy"
    },
    empty_string_branch_is_truthy => {
        "if \"\" then print(\"yes\") else print(\"no\") end\n",
        "yes"
    },
    short_circuit_and_skips_rhs => {
        "local called = false\nlocal function side() called = true return false end\nif false and side() then end\nprint(tostring(called))\n",
        "false"
    },
    short_circuit_or_skips_rhs => {
        "local called = false\nlocal function side() called = true return true end\nif true or side() then end\nprint(tostring(called))\n",
        "false"
    },
    pcall_catches_runtime_error => {
        "local ok = pcall(function() error(\"boom\") end)\nprint(ok)\n",
        "false"
    },
    assert_with_message_fails => {
        "local ok = pcall(function() assert(false, \"nope\") end)\nprint(ok)\n",
        "false"
    },
    metatable_index_fallback_chain => {
        "local defaults = {lang = \"lua\"}\nlocal t = setmetatable({}, {__index = defaults})\nprint(t.lang)\n",
        "lua"
    },
    oop_method_with_colon_passes_self => {
        "local obj = {v = 3}\nfunction obj:double() return self.v * 2 end\nprint(obj:double())\n",
        "6"
    },
    coroutine_producer_consumer_pattern => {
        "local co = coroutine.create(function()\n  coroutine.yield(1)\n  coroutine.yield(2)\nend)\ncoroutine.resume(co)\nlocal _, a = coroutine.resume(co)\nprint(a)\n",
        "2"
    },
    bitwise_mask_extract_nibble => {
        "print((0xAB & 0x0F))\n",
        "11"
    },
    pattern_extract_digits_from_id => {
        "print(string.match(\"user-42\", \"%d+\"))\n",
        "42"
    },
    type_check_before_arithmetic => {
        "local x = \"5\"\nif type(x) == \"number\" then print(x + 1) else print(tonumber(x) + 1) end\n",
        "6"
    },
    nil_safe_table_lookup_with_or => {
        "local t = {}\nlocal v = t.missing or \"default\"\nprint(v)\n",
        "default"
    },
    variadic_sum_with_select => {
        "local function sum(...)\n  local n, s = select(\"#\", ...), 0\n  for i = 1, n do s = s + select(i, ...) end\n  return s\nend\nprint(sum(1, 2, 3, 4))\n",
        "10"
    },
    lexical_scope_shadows_global_name => {
        "x = 1\nlocal function f()\n  local x = 2\n  return x\nend\nprint(f())\n",
        "2"
    },
    module_style_return_table => {
        "local M = {}\nfunction M.add(a, b) return a + b end\nprint(M.add(2, 3))\n",
        "5"
    },
    rotate_array_left_by_one => {
        "local t = {2, 3, 4, 5}\nlocal first = t[1]\nfor i = 1, #t - 1 do t[i] = t[i + 1] end\nt[#t] = first\nprint(t[1])\n",
        "3"
    },
    count_char_occurrences_in_string => {
        "local s, ch, n = \"banana\", \"a\", 0\nfor i = 1, #s do if string.sub(s, i, i) == ch then n = n + 1 end end\nprint(n)\n",
        "3"
    },
    split_on_comma_with_gmatch => {
        "local s = \"a,b,c\"\nlocal out = {}\nfor part in string.gmatch(s, \"[^,]+\") do table.insert(out, part) end\nprint(table.concat(out, \"|\"))\n",
        "a|b|c"
    },
    depth_of_nested_table_walk => {
        "local t = {a = {b = {c = 1}}}\nlocal depth = 0\nlocal node = t\nwhile node.a do depth = depth + 1 node = node.a end\nprint(depth)\n",
        "2"
    },
    gcd_euclidean_algorithm => {
        "local a, b = 48, 18\nwhile b ~= 0 do a, b = b, a % b end\nprint(a)\n",
        "6"
    },
    lcm_from_gcd_formula => {
        "local a, b = 4, 6\nlocal x, y = a, b\nwhile y ~= 0 do x, y = y, x % y end\nprint(a * b / x)\n",
        "12"
    },
    palindrome_check_two_pointers => {
        "local s = \"radar\"\nlocal i, j, ok = 1, #s, true\nwhile i < j do\n  if string.sub(s, i, i) ~= string.sub(s, j, j) then ok = false break end\n  i, j = i + 1, j - 1\nend\nprint(ok)\n",
        "true"
    },
    insertion_sort_small_list => {
        "local t = {5, 2, 4, 1}\nfor i = 2, #t do\n  local key, j = t[i], i - 1\n  while j >= 1 and t[j] > key do t[j + 1] = t[j] j = j - 1 end\n  t[j + 1] = key\nend\nprint(t[1])\n",
        "1"
    },
    dictionary_merge_with_pairs => {
        "local a, b, out = {x = 1}, {y = 2}, {}\nfor k, v in pairs(a) do out[k] = v end\nfor k, v in pairs(b) do out[k] = v end\nprint(out.x + out.y)\n",
        "3"
    },
    flatten_one_level_array => {
        "local nested = {{1, 2}, {3}}\nlocal flat = {}\nfor i = 1, #nested do\n  for j = 1, #nested[i] do table.insert(flat, nested[i][j]) end\nend\nprint(#flat)\n",
        "3"
    },
    zip_two_arrays_into_pairs_table => {
        "local a, b = {1, 2}, {\"x\", \"y\"}\nlocal z = {}\nfor i = 1, math.min(#a, #b) do z[i] = a[i] .. b[i] end\nprint(z[2])\n",
        "2y"
    },
    running_max_stream => {
        "local data = {3, 1, 4, 1, 5}\nlocal max = data[1]\nfor i = 2, #data do if data[i] > max then max = data[i] end end\nprint(max)\n",
        "5"
    },
    simulate_state_machine_two_states => {
        "local state = \"idle\"\nif state == \"idle\" then state = \"run\" end\nif state == \"run\" then state = \"done\" end\nprint(state)\n",
        "done"
    },
    parse_key_value_line => {
        "local line = \"name=vybe\"\nlocal k, v = string.match(line, \"(%w+)=(%w+)\")\nprint(k .. \":\" .. v)\n",
        "name:vybe"
    },
    tokenize_words_with_pattern_gmatch => {
        "local s = \"one two\"\nlocal n = 0\nfor _ in string.gmatch(s, \"%a+\") do n = n + 1 end\nprint(n)\n",
        "2"
    },
}

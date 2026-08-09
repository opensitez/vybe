//! Language core semantics extended tests — assignments, control structures, and evaluation (Lua 5.x)

lua_print! {
assign_multiple_vals => {
    "local a, b = 1, 2\nprint(a .. \",\" .. b)\n",
    "1,2"
},
assign_multiple_swap => {
    "local a, b = 1, 2\na, b = b, a\nprint(a .. \",\" .. b)\n",
    "2,1"
},
assign_too_many_vals => {
    "local a, b = 1, 2, 3\nprint(a .. \",\" .. b)\n",
    "1,2"
},
assign_too_few_vals => {
    "local a, b, c = 1, 2\nprint(a .. \",\" .. b .. \",\" .. tostring(c))\n",
    "1,2,nil"
},
assign_evaluate_first => {
    "local t = {10}\nlocal i = 1\ni, t[i] = 2, 99\nprint(i .. \",\" .. t[1] .. \",\" .. tostring(t[2]))\n",
    "2,99,nil"
},
control_if_true => {
    "local x = false\nif 1 then x = true end\nprint(x)\n",
    "true"
},
control_if_false => {
    "local x = true\nif nil then x = false end\nprint(x)\n",
    "true"
},
control_if_else => {
    "local x\nif false then x = \"if\" else x = \"else\" end\nprint(x)\n",
    "else"
},
control_if_elseif => {
    "local x\nif false then x = \"if\" elseif true then x = \"elseif\" else x = \"else\" end\nprint(x)\n",
    "elseif"
},
control_while_loop => {
    "local n = 0\nwhile n < 3 do n = n + 1 end\nprint(n)\n",
    "3"
},
control_repeat_until => {
    "local n = 0\nrepeat n = n + 1 until n >= 3\nprint(n)\n",
    "3"
} }

-- vybe-test: lua/programs/linked_list_traversal
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "123"
local __i = 0

local function make_node(val, next_node) return {val=val, next=next_node} end
local list = make_node(1, make_node(2, make_node(3, nil)))
local s = ''
local node = list
while node do s = s .. node.val; node = node.next end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end

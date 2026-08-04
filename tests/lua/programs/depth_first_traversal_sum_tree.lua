-- vybe-test: lua/programs/depth_first_traversal_sum_tree
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "10"
local __i = 0

local tree = {v=1, left={v=2}, right={v=3, right={v=4}}}
local function sum(node)
  if node == nil then return 0 end
  return node.v + sum(node.left) + sum(node.right)
end
do local __t = tostring(sum(tree)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end

-- vybe-test: lua/functional_patterns/partial_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "21"
local __i = 0

local function partial(f, ...)
  local args = {...}
  return function(...)
    local all = {table.unpack(args)}
    for _, v in ipairs({...}) do all[#all+1] = v end
    return f(table.unpack(all))
  end
end
local function mul(a, b) return a * b end
local triple = partial(mul, 3)
do local __t = tostring(triple(7)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end

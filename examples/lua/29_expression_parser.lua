-- App 29: Expression Parser + Evaluator (recursive descent)

local input = "(2 + 3) * 4 - 6 / 2 + 7"
local i = 1

local function peek() return input:sub(i, i) end
local function eat_ws() while peek():match("%s") do i = i + 1 end end

local function parse_number()
  eat_ws()
  local s, e = input:find("^%d+", i)
  if not s then error("number expected at " .. i) end
  i = e + 1
  return tonumber(input:sub(s, e))
end

local parse_expr

local function parse_factor()
  eat_ws()
  if peek() == "(" then
    i = i + 1
    local v = parse_expr()
    eat_ws()
    assert(peek() == ")", "missing )")
    i = i + 1
    return v
  end
  return parse_number()
end

local function parse_term()
  local v = parse_factor()
  while true do
    eat_ws()
    local op = peek()
    if op ~= "*" and op ~= "/" then break end
    i = i + 1
    local rhs = parse_factor()
    v = (op == "*") and (v * rhs) or (v / rhs)
  end
  return v
end

parse_expr = function()
  local v = parse_term()
  while true do
    eat_ws()
    local op = peek()
    if op ~= "+" and op ~= "-" then break end
    i = i + 1
    local rhs = parse_term()
    v = (op == "+") and (v + rhs) or (v - rhs)
  end
  return v
end

print(input, "=", parse_expr())

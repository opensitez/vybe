-- 20_mini_test_runner.lua
-- Demonstrates closures + tables + pcall by implementing a tiny test runner.

local tests = {}

local function test(name, fn)
  tests[#tests + 1] = {name = name, fn = fn}
end

local function expect_equal(actual, expected)
  if actual ~= expected then
    error(string.format("expected %s, got %s", tostring(expected), tostring(actual)), 2)
  end
end

test("string.upper works", function()
  expect_equal(string.upper("lua"), "LUA")
end)

test("table insert at end", function()
  local t = {1, 2}
  table.insert(t, 3)
  expect_equal(#t, 3)
  expect_equal(t[3], 3)
end)

test("intentional failure example", function()
  expect_equal(math.floor(2.9), 2)
end)

local passed, failed = 0, 0
for _, t in ipairs(tests) do
  local ok, err = pcall(t.fn)
  if ok then
    passed = passed + 1
    print("PASS", t.name)
  else
    failed = failed + 1
    print("FAIL", t.name, "=>", err)
  end
end

print(string.format("summary: passed=%d failed=%d total=%d", passed, failed, #tests))

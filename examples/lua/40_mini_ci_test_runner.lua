-- App 40: Mini CI Test Runner with retries and report

local tests = {
  {name = "parse_config", flaky = 0.0},
  {name = "db_migration", flaky = 0.4},
  {name = "api_contract", flaky = 0.1},
  {name = "ui_snapshot", flaky = 0.2},
}

local function run_test(test)
  return math.random() >= test.flaky
end

local report = {passed = 0, failed = 0, attempts = 0}

for _, t in ipairs(tests) do
  local ok = false
  for attempt = 1, 2 do
    report.attempts = report.attempts + 1
    ok = run_test(t)
    if ok then break end
  end
  if ok then
    report.passed = report.passed + 1
    print("PASS", t.name)
  else
    report.failed = report.failed + 1
    print("FAIL", t.name)
  end
end

print(string.format("summary passed=%d failed=%d attempts=%d", report.passed, report.failed, report.attempts))

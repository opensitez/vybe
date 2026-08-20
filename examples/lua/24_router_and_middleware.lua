-- App 24: Router and Middleware Engine

local routes = {}
local middleware = {}

local function use(fn)
  middleware[#middleware + 1] = fn
end

local function handle(method, path, fn)
  routes[method .. " " .. path] = fn
end

local function dispatch(req)
  local res = {status = 200, body = ""}
  local index = 1

  local function next_mw()
    if index <= #middleware then
      local fn = middleware[index]
      index = index + 1
      fn(req, res, next_mw)
    else
      local h = routes[req.method .. " " .. req.path]
      if h then h(req, res) else res.status, res.body = 404, "Not Found" end
    end
  end

  next_mw()
  return res
end

use(function(req, res, next)
  req.trace = "req-" .. tostring(math.random(1000, 9999))
  next()
end)

use(function(req, res, next)
  if req.headers.auth ~= "token123" then
    res.status, res.body = 401, "Unauthorized"
    return
  end
  next()
end)

handle("GET", "/health", function(_, res)
  res.body = "ok"
end)

handle("GET", "/users", function(req, res)
  res.body = "users for " .. req.trace
end)

local reqs = {
  {method = "GET", path = "/health", headers = {auth = "bad"}},
  {method = "GET", path = "/users", headers = {auth = "token123"}},
}

for _, req in ipairs(reqs) do
  local r = dispatch(req)
  print(req.method, req.path, r.status, r.body)
end

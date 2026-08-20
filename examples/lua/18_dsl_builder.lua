-- 18_dsl_builder.lua
-- Demonstrates a tiny fluent DSL using callable tables and chaining.

local Query = {}
Query.__index = Query

function Query.new()
  return setmetatable({parts = {}}, Query)
end

function Query:select(...)
  self.parts[#self.parts + 1] = "SELECT " .. table.concat({...}, ", ")
  return self
end

function Query:from(tbl)
  self.parts[#self.parts + 1] = "FROM " .. tbl
  return self
end

function Query:where(clause)
  self.parts[#self.parts + 1] = "WHERE " .. clause
  return self
end

function Query:order_by(expr)
  self.parts[#self.parts + 1] = "ORDER BY " .. expr
  return self
end

function Query:__tostring()
  return table.concat(self.parts, " ")
end

function Query:__call(limit)
  return tostring(self) .. " LIMIT " .. tostring(limit)
end

local q = Query.new()
  :select("id", "name", "score")
  :from("players")
  :where("score >= 1000")
  :order_by("score DESC")

print("query:", tostring(q))
print("query with limit:", q(10))

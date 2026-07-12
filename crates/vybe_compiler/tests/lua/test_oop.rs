//! OOP patterns — `:` vs `.`, methods, constructors (Lua manual §2.5.7, common idiom).

lua_print! {
    colon_syntax_passes_table_as_first_arg => {
        "local t = {}\nfunction t.add(self, x) self.v = (self.v or 0) + x end\nt:add(2)\nprint(t.v)\n",
        "2"
    },
    dot_syntax_call_does_not_pass_self => {
        "local t = {v = 1}\nfunction t.bump(self) self.v = self.v + 1 end\nt.bump(t)\nprint(t.v)\n",
        "2"
    },
    constructor_new_returns_fresh_instance => {
        "local Counter = {}\nfunction Counter.new()\n  return {n = 0}\nend\nlocal c = Counter.new()\nprint(c.n)\n",
        "0"
    },
    method_on_instance_updates_field => {
        "local obj = {value = 1}\nfunction obj:set(v) self.value = v end\nobj:set(5)\nprint(obj.value)\n",
        "5"
    },
    two_instances_do_not_share_fields => {
        "local A = {}\nfunction A.new() return {n = 0} end\nlocal a = A.new()\nlocal b = A.new()\na.n = 3\nprint(b.n)\n",
        "0"
    },
    prototype_lookup_via_index_metatable => {
        "local proto = {greet = function() return \"hi\" end}\nlocal obj = setmetatable({}, {__index = proto})\nprint(obj:greet())\n",
        "hi"
    },
    method_returns_self_for_chaining_pattern => {
        "local Builder = {}\nfunction Builder.new() return setmetatable({}, {__index = Builder}) end\nfunction Builder:add(x)\n  self.parts = (self.parts or \"\") .. x\n  return self\nend\nlocal b = Builder.new():add(\"a\"):add(\"b\")\nprint(b.parts)\n",
        "ab"
    },
    class_table_holds_shared_method => {
        "local Dog = {}\nfunction Dog:speak() return \"woof\" end\nlocal d = setmetatable({}, {__index = Dog})\nprint(d:speak())\n",
        "woof"
    },
    instance_method_reads_own_field_not_proto => {
        "local proto = {name = \"base\"}\nlocal obj = setmetatable({name = \"inst\"}, {__index = proto})\nfunction obj:label() return self.name end\nprint(obj:label())\n",
        "inst"
    },
    static_like_function_on_module_table => {
        "local Math2 = {}\nfunction Math2.twice(x) return x * 2 end\nprint(Math2.twice(4))\n",
        "8"
    },
    colon_call_with_extra_argument => {
        "local t = {}\nfunction t.add(self, a, b) return a + b end\nprint(t:add(2, 3))\n",
        "5"
    },
    inheritance_child_falls_back_to_parent_method => {
        "local Parent = {kind = \"p\"}\nfunction Parent:kind() return self.kind end\nlocal child = setmetatable({}, {__index = Parent})\nprint(child:kind())\n",
        "p"
    },
    colon_vs_dot_self_argument_difference => {
        "local t = {n = 0}\nfunction t.inc(self) self.n = self.n + 1 end\nt:inc()\nprint(t.n)\n",
        "1"
    },
    factory_constructor_sets_initial_state => {
        "local Point = {}\nfunction Point.new(x, y) return {x = x, y = y} end\nlocal p = Point.new(3, 4)\nprint(p.x + p.y)\n",
        "7"
    },
    method_table_on_prototype_not_copied_to_instance => {
        "local Proto = {val = 1}\nlocal a = setmetatable({}, {__index = Proto})\nlocal b = setmetatable({}, {__index = Proto})\na.val = 2\nprint(b.val)\n",
        "1"
    },
    explicit_self_with_dot_call_matches_colon => {
        "local obj = {v = 1}\nfunction obj:set(x) self.v = x end\nobj.set(obj, 9)\nprint(obj.v)\n",
        "9"
    },
    subclass_overrides_parent_method_field => {
        "local Base = {tag = \"b\"}\nfunction Base:tag() return self.tag end\nlocal child = setmetatable({tag = \"c\"}, {__index = Base})\nprint(child:tag())\n",
        "c"
    },
    module_table_namespace_for_functions => {
        "local List = {}\nfunction List.head(t) return t[1] end\nprint(List.head({7, 8}))\n",
        "7"
    },
    instance_method_invokes_with_colon_syntax => {
        "local Acc = {}\nfunction Acc.new() return setmetatable({n = 0}, {__index = Acc}) end\nfunction Acc:add(x) self.n = self.n + x end\nlocal a = Acc.new()\na:add(4)\nprint(a.n)\n",
        "4"
    },
    tostring_metamethod_on_instance => {
        "local Vec = {}\nfunction Vec.new(x, y) return setmetatable({x=x, y=y}, {__tostring = function(v) return '(' .. v.x .. ',' .. v.y .. ')' end}) end\nlocal v = Vec.new(3, 4)\nprint(tostring(v))\n",
        "(3,4)"
    },
    super_call_via_explicit_prototype_reference => {
        "local Animal = {}\nfunction Animal:sound() return 'generic' end\nlocal Dog = setmetatable({}, {__index = Animal})\nfunction Dog:sound() return Animal.sound(self) .. '+woof' end\nlocal d = setmetatable({}, {__index = Dog})\nprint(d:sound())\n",
        "generic+woof"
    },
    private_state_via_closure_in_constructor => {
        "local function make_counter(init)\n  local n = init\n  return {\n    get = function() return n end,\n    inc = function() n = n + 1 end,\n  }\nend\nlocal c = make_counter(10)\nc.inc(); c.inc()\nprint(c.get())\n",
        "12"
    },
    mixin_copies_methods_into_target_class => {
        "local function mixin(target, source)\n  for k, v in pairs(source) do target[k] = v end\nend\nlocal Fly = {fly = function(self) return self.name .. ' flies' end}\nlocal Bird = {name = 'eagle'}\nmixin(Bird, Fly)\nprint(Bird:fly())\n",
        "eagle flies"
    },
    instanceof_check_via_metatable_comparison => {
        "local MyClass = {}\nMyClass.__index = MyClass\nfunction MyClass.new() return setmetatable({}, MyClass) end\nlocal obj = MyClass.new()\nprint(getmetatable(obj) == MyClass)\n",
        "true"
    },
    index_as_function_for_computed_lookup => {
        "local t = setmetatable({}, {\n  __index = function(tbl, key)\n    return key:upper()\n  end\n})\nprint(t.hello .. t.world)\n",
        "HELLOWORLD"
    },
    method_override_in_subclass_shadows_parent => {
        "local Base = {}\nBase.__index = Base\nfunction Base:describe() return 'base' end\nlocal Child = setmetatable({}, {__index = Base})\nChild.__index = Child\nfunction Child:describe() return 'child' end\nlocal obj = setmetatable({}, Child)\nprint(obj:describe())\n",
        "child"
    },
    constructor_collecting_varargs_into_field => {
        "local function make_list(...)\n  return {items = {...}}\nend\nlocal l = make_list(10, 20, 30)\nprint(#l.items .. ',' .. l.items[2])\n",
        "3,20"
    },
}

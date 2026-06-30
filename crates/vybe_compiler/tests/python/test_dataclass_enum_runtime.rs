//! dataclasses, Enum, TypedDict, NamedTuple runtime.

crate::runtime_case!(
    dataclass_basic,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\n y: int = 0\nprint(P(1).x)\n",
    "1"
);
crate::runtime_case!(
    dataclass_default,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int = 5\nprint(P().x)\n",
    "5"
);
crate::runtime_case!(
    dataclass_eq,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\nprint(P(1) == P(1))\n",
    "True"
);
crate::runtime_case!(
    dataclass_ne,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\nprint(P(1) != P(2))\n",
    "True"
);
crate::runtime_case!(
    dataclass_repr,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\nprint('P' in repr(P(1)))\n",
    "True"
);
crate::runtime_case!(
    dataclass_asdict,
    "from dataclasses import dataclass, asdict\n@dataclass\nclass P:\n x: int\nprint(asdict(P(1)))\n",
    "{'x': 1}"
);
crate::runtime_case!(
    dataclass_astuple,
    "from dataclasses import dataclass, astuple\n@dataclass\nclass P:\n x: int\n y: int\nprint(astuple(P(1, 2)))\n",
    "(1, 2)"
);
crate::runtime_case!(
    dataclass_replace,
    "from dataclasses import dataclass, replace\n@dataclass\nclass P:\n x: int\nprint(replace(P(1), x=9).x)\n",
    "9"
);
crate::runtime_case!(
    dataclass_fields,
    "from dataclasses import dataclass, fields\n@dataclass\nclass P:\n x: int\nprint(len(fields(P(1))))\n",
    "1"
);
crate::runtime_case!(
    dataclass_frozen,
    "from dataclasses import dataclass\n@dataclass(frozen=True)\nclass P:\n x: int\nprint(P(1).x)\n",
    "1"
);
crate::runtime_case!(
    dataclass_order_false,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\nprint(P(1).x)\n",
    "1"
);
crate::runtime_case!(
    enum_basic,
    "from enum import Enum\nclass Color(Enum):\n RED = 1\n GREEN = 2\nprint(Color.RED.value)\n",
    "1"
);
crate::runtime_case!(
    enum_name,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(E.A.name)\n",
    "A"
);
crate::runtime_case!(
    enum_member,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(E.A is E.A)\n",
    "True"
);
crate::runtime_case!(
    enum_iteration,
    "from enum import Enum\nclass E(Enum):\n A = 1\n B = 2\nprint(len(list(E)))\n",
    "2"
);
crate::runtime_case!(
    enum_auto,
    "from enum import Enum, auto\nclass E(Enum):\n A = auto()\n B = auto()\nprint(E.B.value > E.A.value)\n",
    "True"
);
crate::runtime_case!(
    intenum_value,
    "from enum import IntEnum\nclass E(IntEnum):\n X = 1\nprint(E.X + 1)\n",
    "2"
);
crate::runtime_case!(
    flag_or,
    "from enum import Flag, auto\nclass F(Flag):\n A = auto()\n B = auto()\nprint((F.A | F.B).name)\n",
    "A|B"
);
crate::runtime_case!(
    enum_str,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(str(E.A))\n",
    "E.A"
);
crate::runtime_case!(
    enum_repr,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(repr(E.A))\n",
    "<E.A: 1>"
);
crate::runtime_case!(
    namedtuple_field,
    "from collections import namedtuple\nP = namedtuple('P', 'x y')\nprint(P(1, 2).y)\n",
    "2"
);
crate::runtime_case!(
    namedtuple_asdict,
    "from collections import namedtuple\nP = namedtuple('P', 'x')\nprint(P(1)._asdict())\n",
    "{'x': 1}"
);
crate::runtime_case!(
    namedtuple_replace,
    "from collections import namedtuple\nP = namedtuple('P', 'x')\nprint(P(1)._replace(x=9).x)\n",
    "9"
);
crate::runtime_case!(
    namedtuple_len,
    "from collections import namedtuple\nP = namedtuple('P', 'a b c')\nprint(len(P(1, 2, 3)))\n",
    "3"
);
crate::runtime_case!(
    typeddict_runtime,
    "from typing import TypedDict\nclass D(TypedDict):\n x: int\nprint(D(x=1)['x'])\n",
    "1"
);
crate::runtime_case!(
    dataclass_post_init,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\n def __post_init__(self):\n  self.y = self.x + 1\nprint(P(1).y)\n",
    "2"
);
crate::runtime_case!(
    dataclass_init_var,
    "from dataclasses import dataclass, InitVar\n@dataclass\nclass P:\n x: int\n iv: InitVar[int]\n def __post_init__(self, iv):\n  self.x += iv\nprint(P(1, 2).x)\n",
    "3"
);
crate::runtime_case!(
    enum_unique_values,
    "from enum import Enum\nclass E(Enum):\n A = 1\n B = 2\nprint(E.A.value)\n",
    "1"
);
crate::runtime_case!(
    enum_comparison,
    "from enum import Enum\nclass E(Enum):\n A = 1\n B = 2\nprint(E.A != E.B)\n",
    "True"
);
crate::runtime_case!(
    dataclass_hash_default,
    "from dataclasses import dataclass\n@dataclass\nclass P:\n x: int\nprint(hash(P(1)) == hash(P(1)))\n",
    "True"
);
crate::runtime_case!(
    enum_has_value,
    "from enum import Enum\nclass E(Enum):\n A = 'x'\nprint(E.A.value)\n",
    "x"
);
crate::runtime_case!(
    dataclass_field_factory,
    "from dataclasses import dataclass, field\n@dataclass\nclass P:\n xs: list = field(default_factory=list)\nprint(P().xs)\n",
    "[]"
);
crate::runtime_case!(
    enum_membership,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(E.A in E)\n",
    "True"
);
crate::runtime_case!(
    namedtuple_index,
    "from collections import namedtuple\nP = namedtuple('P', 'x y')\nprint(P(1, 2)[0])\n",
    "1"
);
crate::runtime_case!(
    namedtuple_iter,
    "from collections import namedtuple\nP = namedtuple('P', 'x y')\nprint(list(P(1, 2)))\n",
    "[1, 2]"
);
crate::runtime_case!(
    enum_by_name,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(E['A'].value)\n",
    "1"
);
crate::runtime_case!(
    enum_by_value,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(E(1).name)\n",
    "A"
);
crate::runtime_case!(
    dataclass_slots,
    "from dataclasses import dataclass\n@dataclass(slots=True)\nclass P:\n x: int\nprint(P(1).x)\n",
    "1"
);
crate::runtime_case!(
    flag_contains,
    "from enum import Flag, auto\nclass F(Flag):\n A = auto()\n B = auto()\nprint(F.A in (F.A | F.B))\n",
    "True"
);
crate::runtime_case!(
    intflag_int_compare,
    "from enum import IntEnum\nclass E(IntEnum):\n X = 1\nprint(E.X == 1)\n",
    "True"
);
crate::runtime_case!(
    dataclass_is_dataclass,
    "from dataclasses import dataclass, is_dataclass\n@dataclass\nclass P:\n x: int\nprint(is_dataclass(P))\n",
    "True"
);
crate::runtime_case!(
    dataclass_not_dataclass,
    "from dataclasses import is_dataclass\nclass C:\n pass\nprint(is_dataclass(C))\n",
    "False"
);
crate::runtime_case!(
    enum_mixed_inheritance,
    "from enum import Enum\nclass E(str, Enum):\n A = 'a'\nprint(E.A + 'b')\n",
    "ab"
);
crate::runtime_case!(
    namedtuple_defaults,
    "from collections import namedtuple\nP = namedtuple('P', 'x y', defaults=[0])\nprint(P(1).y)\n",
    "0"
);
crate::runtime_case!(
    typeddict_total,
    "from typing import TypedDict\nclass D(TypedDict, total=False):\n x: int\nprint(D().get('x') is None)\n",
    "True"
);
crate::runtime_case!(
    enum_bool,
    "from enum import Enum\nclass E(Enum):\n A = 1\nprint(bool(E.A))\n",
    "True"
);

crate::compile_case!(dataclass_kw_only, "from dataclasses import dataclass, KW_ONLY\n@dataclass(kw_only=True)\nclass P:\n x: int\n");
crate::compile_case!(enum_unique_decorator, "from enum import Enum, unique\n@unique\nclass E(Enum):\n A = 1\n");
crate::compile_case!(typeddict_required_notrequired, "from typing import TypedDict, Required, NotRequired\nclass D(TypedDict):\n x: Required[int]\n");
crate::compile_case!(protocol_runtime, "from typing import Protocol\nclass P(Protocol):\n def m(self) -> int: ...\n");
crate::compile_case!(enum_strenum, "from enum import StrEnum\nclass E(StrEnum):\n A = 'a'\n");

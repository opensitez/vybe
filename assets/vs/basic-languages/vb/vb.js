/*!-----------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Version: 0.50.0(c321d0fbecb50ab8a5365fa1965476b0ae63fc87)
 * Released under the MIT license
 * https://github.com/microsoft/monaco-editor/blob/main/LICENSE.txt
 * Modified for vybe VB6/VB.NET support
 *-----------------------------------------------------------------------------*/
define("vs/basic-languages/vb/vb", ["require","require"], (require) => {
"use strict";

var conf = {
    comments: {
        lineComment: "'",
        blockComment: ["/*", "*/"]
    },
    brackets: [
        ["{", "}"],
        ["[", "]"],
        ["(", ")"],
        ["<", ">"],
        ["addhandler", "end addhandler"],
        ["class", "end class"],
        ["enum", "end enum"],
        ["event", "end event"],
        ["function", "end function"],
        ["get", "end get"],
        ["if", "end if"],
        ["interface", "end interface"],
        ["module", "end module"],
        ["namespace", "end namespace"],
        ["operator", "end operator"],
        ["property", "end property"],
        ["raiseevent", "end raiseevent"],
        ["removehandler", "end removehandler"],
        ["select", "end select"],
        ["set", "end set"],
        ["structure", "end structure"],
        ["sub", "end sub"],
        ["synclock", "end synclock"],
        ["try", "end try"],
        ["while", "end while"],
        ["with", "end with"],
        ["using", "end using"],
        ["do", "loop"],
        ["for", "next"]
    ],
    autoClosingPairs: [
        { open: "{", close: "}", notIn: ["string", "comment"] },
        { open: "[", close: "]", notIn: ["string", "comment"] },
        { open: "(", close: ")", notIn: ["string", "comment"] },
        { open: '"', close: '"', notIn: ["string", "comment"] },
        { open: "<", close: ">", notIn: ["string", "comment"] }
    ],
    folding: {
        markers: {
            start: new RegExp("^\\s*#Region\\b"),
            end: new RegExp("^\\s*#End Region\\b")
        }
    }
};

var language = {
    defaultToken: "",
    tokenPostfix: ".vb",
    ignoreCase: true,
    
    brackets: [
        { token: "delimiter.bracket", open: "{", close: "}" },
        { token: "delimiter.array", open: "[", close: "]" },
        { token: "delimiter.parenthesis", open: "(", close: ")" },
        { token: "delimiter.angle", open: "<", close: ">" },
        { token: "keyword.tag-addhandler", open: "addhandler", close: "end addhandler" },
        { token: "keyword.tag-class", open: "class", close: "end class" },
        { token: "keyword.tag-enum", open: "enum", close: "end enum" },
        { token: "keyword.tag-event", open: "event", close: "end event" },
        { token: "keyword.tag-function", open: "function", close: "end function" },
        { token: "keyword.tag-get", open: "get", close: "end get" },
        { token: "keyword.tag-if", open: "if", close: "end if" },
        { token: "keyword.tag-interface", open: "interface", close: "end interface" },
        { token: "keyword.tag-module", open: "module", close: "end module" },
        { token: "keyword.tag-namespace", open: "namespace", close: "end namespace" },
        { token: "keyword.tag-operator", open: "operator", close: "end operator" },
        { token: "keyword.tag-property", open: "property", close: "end property" },
        { token: "keyword.tag-raiseevent", open: "raiseevent", close: "end raiseevent" },
        { token: "keyword.tag-removehandler", open: "removehandler", close: "end removehandler" },
        { token: "keyword.tag-select", open: "select", close: "end select" },
        { token: "keyword.tag-set", open: "set", close: "end set" },
        { token: "keyword.tag-structure", open: "structure", close: "end structure" },
        { token: "keyword.tag-sub", open: "sub", close: "end sub" },
        { token: "keyword.tag-synclock", open: "synclock", close: "end synclock" },
        { token: "keyword.tag-try", open: "try", close: "end try" },
        { token: "keyword.tag-while", open: "while", close: "end while" },
        { token: "keyword.tag-with", open: "with", close: "end with" },
        { token: "keyword.tag-using", open: "using", close: "end using" },
        { token: "keyword.tag-do", open: "do", close: "loop" },
        { token: "keyword.tag-for", open: "for", close: "next" }
    ],
    
    keywords: [
        // Control flow
        "If", "Then", "Else", "ElseIf", "End", "EndIf", "Select", "Case", 
        "For", "To", "Step", "Next", "Each", "In", "While", "Wend", "Do", "Loop", "Until",
        "With", "Try", "Catch", "Finally", "When", "Exit", "Continue", "Return",
        
        // Declarations
        "Dim", "ReDim", "Preserve", "Const", "As", "New", "Set",
        "Sub", "Function", "Property", "Get", "Module", "Class", "Enum", 
        "Inherits", "Implements", "Handles",
        
        // Modifiers
        "Public", "Private", "Protected", "Friend", "Shared", "Static",
        "Partial", "MustOverride", "MustInherit", "NotOverridable", "NotInheritable",
        "Overrides", "Overloads", "Overridable", "Shadows", "WithEvents",
        "ReadOnly", "WriteOnly", "Optional", "ByVal", "ByRef", "ParamArray",
        
        // Operators & Logic
        "And", "AndAlso", "Or", "OrElse", "Not", "Mod", "Xor", 
        "Is", "IsNot", "Like", "TypeOf",
        
        // Literals & Constants
        "True", "False", "Nothing", "Me", "MyBase", "MyClass",
        
        // Types
        "Boolean", "Byte", "SByte", "Short", "UShort", "Integer", "UInteger",
        "Long", "ULong", "Single", "Double", "Decimal", "Char", "String", 
        "Date", "Object", "Variant",
        
        // Type operations
        "Of", "DirectCast", "TryCast", "CType", "GetType",
        "CBool", "CByte", "CChar", "CDate", "CDbl", "CDec", "CInt", "CLng", 
        "CObj", "CSByte", "CShort", "CSng", "CStr", "CUInt", "CULng", "CUShort",
        
        // VB6 File I/O
        "Open", "Close", "Print", "Line", "Input", "Write", "Append", "Output", 
        "Binary", "Random",
        
        // Other keywords
        "Call", "Await", "Async", "Imports", "Namespace", "Interface", "Structure",
        "Event", "Delegate", "AddHandler", "RemoveHandler", "RaiseEvent",
        "Operator", "Widening", "Narrowing", "Using", "SyncLock",
        "Option", "Declare", "Lib", "Alias", "Erase", "Error",
        "GoTo", "GoSub", "On", "Resume", "Stop", "Let",
        "Default", "Global", "Out", "NameOf", "GetXMLNamespace"
    ],
    
    tagwords: [
        "If", "Sub", "Select", "Try", "Class", "Enum", "Function", "Get", 
        "Interface", "Module", "Namespace", "Operator", "Set", "Structure", 
        "Using", "While", "With", "Do", "Loop", "For", "Next", "Property", 
        "Continue", "AddHandler", "RemoveHandler", "Event", "RaiseEvent", "SyncLock"
    ],
    
    symbols: /[=><!~?;\.,:&|+\-*\/\^%]+/,
    integersuffix: /U?[DI%L&S@]?/,
    floatsuffix: /[R#F!]?/,
    
    tokenizer: {
        root: [
            { include: "@whitespace" },
            
            // Special handling for "End" keyword combinations
            [/end\s+if/i, "keyword"],
            [/end\s+sub/i, "keyword"],
            [/end\s+function/i, "keyword"],
            [/end\s+class/i, "keyword"],
            [/end\s+module/i, "keyword"],
            [/end\s+namespace/i, "keyword"],
            [/end\s+enum/i, "keyword"],
            [/end\s+interface/i, "keyword"],
            [/end\s+property/i, "keyword"],
            [/end\s+structure/i, "keyword"],
            [/end\s+select/i, "keyword"],
            [/end\s+try/i, "keyword"],
            [/end\s+while/i, "keyword"],
            [/end\s+with/i, "keyword"],
            [/end\s+using/i, "keyword"],
            [/end\s+get/i, "keyword"],
            [/end\s+set/i, "keyword"],
            [/end\s+operator/i, "keyword"],
            [/end\s+event/i, "keyword"],
            [/end\s+addhandler/i, "keyword"],
            [/end\s+removehandler/i, "keyword"],
            [/end\s+raiseevent/i, "keyword"],
            [/end\s+synclock/i, "keyword"],
            
            // Keywords
            [/[a-zA-Z_]\w*/, {
                cases: {
                    "@tagwords": { token: "keyword.tag-$0" },
                    "@keywords": { token: "keyword" },
                    "@default": "identifier"
                }
            }],
            
            // Preprocessor directives
            [/^\s*#\w+/, "keyword.directive"],
            
            // Numbers
            [/\d*\d+e([\-+]?\d+)?(@floatsuffix)/, "number.float"],
            [/\d*\.\d+(e[\-+]?\d+)?(@floatsuffix)/, "number.float"],
            [/&H[0-9a-f]+(@integersuffix)/i, "number.hex"],
            [/&O[0-7]+(@integersuffix)/i, "number.octal"],
            [/\d+(@integersuffix)/, "number"],
            [/#.*#/, "number.date"],
            
            // Delimiters
            [/[{}()\[\]]/, "@brackets"],
            [/@symbols/, "delimiter"],
            
            // Strings
            [/["\u201c\u201d]/, { token: "string.quote", next: "@string" }]
        ],
        
        whitespace: [
            [/[ \t\r\n]+/, ""],
            [/(\'|REM(?!\w)).*$/, "comment"]
        ],
        
        string: [
            [/[^"\u201c\u201d]+/, "string"],
            [/["\u201c\u201d]{2}/, "string.escape"],
            [/["\u201c\u201d]C?/, { token: "string.quote", next: "@pop" }]
        ]
    }
};

return { conf: conf, language: language };
});

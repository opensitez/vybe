use super::{Expression, Identifier, Statement, VBType};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Declaration {
    Variable(VariableDecl),
    Constant(ConstDecl),
    Sub(SubDecl),
    Function(FunctionDecl),
    Class(ClassDecl),
    Enum(EnumDecl),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct VariableDecl {
    pub name: Identifier,
    pub var_type: Option<VBType>,
    pub array_bounds: Option<Vec<Expression>>,
    pub initializer: Option<Expression>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ConstDecl {
    pub visibility: Visibility,
    pub name: Identifier,
    pub const_type: VBType,
    pub value: Expression,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SubDecl {
    pub visibility: Visibility,
    pub name: Identifier,
    pub parameters: Vec<Parameter>,
    pub body: Vec<Statement>,
    pub handles: Option<Vec<String>>,
    #[serde(default)]
    pub is_async: bool,
    #[serde(default)]
    pub is_extension: bool,
    #[serde(default)]
    pub is_overridable: bool,
    #[serde(default)]
    pub is_overrides: bool,
    #[serde(default)]
    pub is_must_override: bool,
    #[serde(default)]
    pub is_shared: bool,
    #[serde(default)]
    pub is_not_overridable: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FunctionDecl {
    pub visibility: Visibility,
    pub name: Identifier,
    pub parameters: Vec<Parameter>,
    pub return_type: Option<VBType>,
    pub body: Vec<Statement>,
    #[serde(default)]
    pub is_async: bool,
    #[serde(default)]
    pub is_extension: bool,
    #[serde(default)]
    pub is_overridable: bool,
    #[serde(default)]
    pub is_overrides: bool,
    #[serde(default)]
    pub is_must_override: bool,
    #[serde(default)]
    pub is_shared: bool,
    #[serde(default)]
    pub is_not_overridable: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ClassDecl {
    pub visibility: Visibility,
    pub name: Identifier,
    pub is_partial: bool,
    pub inherits: Option<VBType>,
    pub implements: Vec<VBType>,
    pub properties: Vec<PropertyDecl>,
    pub methods: Vec<MethodDecl>,
    pub fields: Vec<VariableDecl>,
    #[serde(default)]
    pub is_must_inherit: bool,
    #[serde(default)]
    pub is_not_inheritable: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PropertyDecl {
    pub visibility: Visibility,
    pub name: Identifier,
    pub parameters: Vec<Parameter>,
    pub return_type: Option<VBType>,
    pub getter: Option<Vec<Statement>>,
    pub setter: Option<(Parameter, Vec<Statement>)>, // Setter has a value parameter and a body
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum MethodDecl {
    Sub(SubDecl),
    Function(FunctionDecl),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Parameter {
    pub pass_type: ParameterPassType,
    pub name: Identifier,
    pub param_type: Option<VBType>,
    #[serde(default)]
    pub is_optional: bool,
    #[serde(default)]
    pub default_value: Option<Expression>,
    #[serde(default)]
    pub is_nullable: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum ParameterPassType {
    ByVal,
    ByRef,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum Visibility {
    Public,
    Private,
    Protected,
    Friend,
}

impl Default for Visibility {
    fn default() -> Self {
        Visibility::Public
    }
}

impl Default for ParameterPassType {
    fn default() -> Self {
        ParameterPassType::ByRef
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct EnumDecl {
    pub visibility: Visibility,
    pub name: Identifier,
    pub members: Vec<EnumMember>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct EnumMember {
    pub name: Identifier,
    pub value: Option<Expression>,
}

//! `nameof` expressions on types, variables, methods, and members.
//! GAP: nameof coverage is thin in the existing suite.

use crate::csharp_cases;

csharp_cases! {
    nameof_local_int_variable_returns_identifier => {
        r#"int itemCount=0; Console.WriteLine(nameof(itemCount));"#,
        ["itemCount"]
    };

    nameof_local_string_variable_returns_identifier => {
        r#"string label="x"; Console.WriteLine(nameof(label));"#,
        ["label"]
    };

    nameof_local_bool_variable_returns_identifier => {
        r#"bool isReady=true; Console.WriteLine(nameof(isReady));"#,
        ["isReady"]
    };

    nameof_method_parameter_returns_parameter_name => {
        r#"void Report(int total){Console.WriteLine(nameof(total));} Report(1);"#,
        ["total"]
    };

    nameof_two_parameters_print_both_names => {
        r#"void Pair(int left,int right){Console.WriteLine(nameof(left)); Console.WriteLine(nameof(right));} Pair(1,2);"#,
        ["left", "right"]
    };

    nameof_static_method_on_type_returns_method_name => {
        r#"class MathUtil{public static int Double(int n)=>n*2;} Console.WriteLine(nameof(MathUtil.Double));"#,
        ["Double"]
    };

    nameof_instance_method_on_type_returns_method_name => {
        r#"class Worker{public void Run(){}} Console.WriteLine(nameof(Worker.Run));"#,
        ["Run"]
    };

    nameof_local_function_returns_function_name => {
        r#"int Compute(){return 1;} Console.WriteLine(nameof(Compute));"#,
        ["Compute"]
    };

    nameof_public_field_on_type_returns_field_name => {
        r#"class Account{public int Balance;} Console.WriteLine(nameof(Account.Balance));"#,
        ["Balance"]
    };

    nameof_private_field_on_type_returns_field_name => {
        r#"class Vault{private string Secret;} Console.WriteLine(nameof(Vault.Secret));"#,
        ["Secret"]
    };

    nameof_property_getter_member_returns_property_name => {
        r#"class Person{public string Name{get;set;}} Console.WriteLine(nameof(Person.Name));"#,
        ["Name"]
    };

    nameof_static_property_returns_property_name => {
        r#"class Config{public static int Port{get;set;}=80;} Console.WriteLine(nameof(Config.Port));"#,
        ["Port"]
    };

    nameof_event_member_returns_event_name => {
        r#"class Publisher{public event System.Action Raised;} Console.WriteLine(nameof(Publisher.Raised));"#,
        ["Raised"]
    };

    nameof_nested_type_member_returns_member_name => {
        r#"class Outer{public class Inner{public int Value;}} Console.WriteLine(nameof(Outer.Inner.Value));"#,
        ["Value"]
    };

    nameof_nested_type_returns_type_name => {
        r#"class Outer{public class Inner{}} Console.WriteLine(nameof(Outer.Inner));"#,
        ["Inner"]
    };

    nameof_enum_type_returns_type_name => {
        r#"enum Color{Red,Green,Blue} Console.WriteLine(nameof(Color));"#,
        ["Color"]
    };

    nameof_enum_member_returns_member_name => {
        r#"enum Color{Red,Green,Blue} Console.WriteLine(nameof(Color.Green));"#,
        ["Green"]
    };

    nameof_delegate_type_returns_type_name => {
        r#"delegate int Transformer(int value); Console.WriteLine(nameof(Transformer));"#,
        ["Transformer"]
    };

    nameof_interface_type_returns_type_name => {
        r#"interface IRepository{} Console.WriteLine(nameof(IRepository));"#,
        ["IRepository"]
    };

    nameof_struct_type_returns_type_name => {
        r#"struct Point{public int X;} Console.WriteLine(nameof(Point));"#,
        ["Point"]
    };

    nameof_record_type_returns_type_name => {
        r#"record Book(string Title,int Pages); Console.WriteLine(nameof(Book));"#,
        ["Book"]
    };

    nameof_record_primary_constructor_parameter => {
        r#"record Book(string Title,int Pages); Console.WriteLine(nameof(Book.Title));"#,
        ["Title"]
    };

    nameof_generic_type_definition_returns_type_name => {
        r#"class Box<T>{public T Item;} Console.WriteLine(nameof(Box));"#,
        ["Box"]
    };

    nameof_closed_generic_type_returns_type_name => {
        r#"class Box<T>{public T Item;} Console.WriteLine(nameof(Box<int>));"#,
        ["Box"]
    };

    nameof_generic_type_member_returns_member_name => {
        r#"class Box<T>{public T Item;} Console.WriteLine(nameof(Box<string>.Item));"#,
        ["Item"]
    };

    nameof_namespace_qualified_type_returns_simple_name => {
        r#"Console.WriteLine(nameof(System.String));"#,
        ["String"]
    };

    nameof_bcl_type_returns_type_name => {
        r#"Console.WriteLine(nameof(System.DateTime));"#,
        ["DateTime"]
    };

    nameof_console_type_returns_type_name => {
        r#"Console.WriteLine(nameof(System.Console));"#,
        ["Console"]
    };

    nameof_typeof_int_expression_returns_keyword_name => {
        r#"Console.WriteLine(nameof(int));"#,
        ["int"]
    };

    nameof_typeof_string_expression_returns_keyword_name => {
        r#"Console.WriteLine(nameof(string));"#,
        ["string"]
    };

    nameof_typeof_bool_expression_returns_keyword_name => {
        r#"Console.WriteLine(nameof(bool));"#,
        ["bool"]
    };

    nameof_in_string_concatenation_produces_literal_fragment => {
        r#"int age=30; Console.WriteLine("field="+nameof(age));"#,
        ["field=age"]
    };

    nameof_in_interpolated_string_embeds_name => {
        r#"string title="demo"; Console.WriteLine($"name={nameof(title)}");"#,
        ["name=title"]
    };

    nameof_of_method_group_in_expression_context => {
        r#"class Ops{public void Execute(){}} Console.WriteLine(nameof(Ops)+"."+nameof(Ops.Execute));"#,
        ["Ops.Execute"]
    };

    nameof_parameter_in_local_function => {
        r#"void Outer(){void Inner(int offset){Console.WriteLine(nameof(offset));} Inner(3);} Outer();"#,
        ["offset"]
    };

    nameof_static_field_returns_field_name => {
        r#"class Counter{public static int Total=0;} Console.WriteLine(nameof(Counter.Total));"#,
        ["Total"]
    };

    nameof_const_field_returns_field_name => {
        r#"class Limits{public const int Max=100;} Console.WriteLine(nameof(Limits.Max));"#,
        ["Max"]
    };

    nameof_readonly_field_returns_field_name => {
        r#"class Token{public readonly string Value="x";} Console.WriteLine(nameof(Token.Value));"#,
        ["Value"]
    };

    nameof_indexer_declaring_type_returns_type_name => {
        r#"class Bag{public int this[int i]{get=>i;set{}}} Console.WriteLine(nameof(Bag));"#,
        ["Bag"]
    };

    nameof_extension_method_target_type_member => {
        r#"static class Extensions{public static int Twice(this int n)=>n*2;} Console.WriteLine(nameof(Extensions.Twice));"#,
        ["Twice"]
    };

    nameof_overloaded_method_uses_simple_name => {
        r#"class Calc{public int Add(int a,int b)=>a+b; public double Add(double a,double b)=>a+b;} Console.WriteLine(nameof(Calc.Add));"#,
        ["Add"]
    };

    nameof_operator_method_returns_method_name => {
        r#"class Vector{public static Vector operator +(Vector a,Vector b)=>a; public int X;} Console.WriteLine(nameof(Vector.op_Addition));"#,
        ["op_Addition"]
    };

    nameof_partial_class_member_returns_member_name => {
        r#"partial class Partial{public int Id;} Console.WriteLine(nameof(Partial.Id));"#,
        ["Id"]
    };

    nameof_qualified_alias_target_type => {
        r#"using Text=System.String; Console.WriteLine(nameof(Text));"#,
        ["Text"]
    };

    nameof_var_inferred_local_returns_identifier => {
        r#"var delta=1; Console.WriteLine(nameof(delta));"#,
        ["delta"]
    };

    nameof_catch_exception_variable_when_supported => {
        r#"try{throw new System.Exception("x");}catch(System.Exception ex){Console.WriteLine(nameof(ex));}"#,
        ["ex"]
    };

    nameof_foreach_iteration_variable_in_local_function => {
        r#"void Scan(){foreach(var entry in new string[]{"a"}){Console.WriteLine(nameof(entry)); break;}} Scan();"#,
        ["entry"]
    };

    nameof_type_parameter_name_on_generic_method => {
        r#"class Factory{public T Build<T>(T value)=>value;} Console.WriteLine(nameof(Factory.Build));"#,
        ["Build"]
    };

    nameof_multiple_locals_in_one_statement_list => {
        r#"int width=1,height=2; Console.WriteLine(nameof(width)); Console.WriteLine(nameof(height));"#,
        ["width", "height"]
    };
}

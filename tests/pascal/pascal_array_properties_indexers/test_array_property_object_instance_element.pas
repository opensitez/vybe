// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_object_instance_element
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TNode = class public LabelStr: String; constructor Create(L: String); end;
type TNodeMap = class
  private FNodes: array[0..1] of TNode;
  private function GetNode(i: Integer): TNode;
  private procedure SetNode(i: Integer; n: TNode);
  public property Nodes[i: Integer]: TNode read GetNode write SetNode; default;
  public destructor Destroy; override;
end;
constructor TNode.Create(L: String); begin LabelStr := L; end;
function TNodeMap.GetNode(i: Integer): TNode; begin Result := FNodes[i]; end;
procedure TNodeMap.SetNode(i: Integer; n: TNode); begin FNodes[i] := n; end;
destructor TNodeMap.Destroy; begin FNodes[0].Free; FNodes[1].Free; inherited Destroy; end;
var nm: TNodeMap;
begin
  nm := TNodeMap.Create;
  nm[0] := TNode.Create('Start');
  nm[1] := TNode.Create('End');
  WriteLn(nm[0].LabelStr + '-' + nm[1].LabelStr);
  nm.Free;
end.

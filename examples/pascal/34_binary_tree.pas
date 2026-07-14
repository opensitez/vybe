program BinaryTreeDemo;

type
  PTreeNode = ^TTreeNode;
  TTreeNode = record
    Value: Integer;
    Left, Right: PTreeNode;
  end;

  TBinaryTree = class
  private
    FRoot: PTreeNode;
    procedure InsertNode(var Node: PTreeNode; Value: Integer);
    function SearchNode(Node: PTreeNode; Value: Integer): Boolean;
    procedure InOrderTraverse(Node: PTreeNode);
    function CountNodes(Node: PTreeNode): Integer;
    function TreeHeight(Node: PTreeNode): Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Insert(Value: Integer);
    function Contains(Value: Integer): Boolean;
    procedure PrintInOrder;
    function Count: Integer;
    function Height: Integer;
  end;

constructor TBinaryTree.Create;
begin
  FRoot := nil;
end;

procedure FreeTree(Node: PTreeNode);
begin
  if Node = nil then Exit;
  FreeTree(Node^.Left);
  FreeTree(Node^.Right);
  Dispose(Node);
end;

destructor TBinaryTree.Destroy;
begin
  FreeTree(FRoot);
end;

procedure TBinaryTree.InsertNode(var Node: PTreeNode; Value: Integer);
begin
  if Node = nil then
  begin
    New(Node);
    Node^.Value := Value;
    Node^.Left := nil;
    Node^.Right := nil;
  end
  else if Value < Node^.Value then
    InsertNode(Node^.Left, Value)
  else
    InsertNode(Node^.Right, Value);
end;

procedure TBinaryTree.Insert(Value: Integer);
begin
  InsertNode(FRoot, Value);
end;

function TBinaryTree.SearchNode(Node: PTreeNode; Value: Integer): Boolean;
begin
  if Node = nil then
    Result := False
  else if Node^.Value = Value then
    Result := True
  else if Value < Node^.Value then
    Result := SearchNode(Node^.Left, Value)
  else
    Result := SearchNode(Node^.Right, Value);
end;

function TBinaryTree.Contains(Value: Integer): Boolean;
begin
  Result := SearchNode(FRoot, Value);
end;

procedure TBinaryTree.InOrderTraverse(Node: PTreeNode);
begin
  if Node = nil then Exit;
  InOrderTraverse(Node^.Left);
  Write(Node^.Value, ' ');
  InOrderTraverse(Node^.Right);
end;

procedure TBinaryTree.PrintInOrder;
begin
  InOrderTraverse(FRoot);
  Writeln;
end;

function TBinaryTree.CountNodes(Node: PTreeNode): Integer;
begin
  if Node = nil then
    Result := 0
  else
    Result := 1 + CountNodes(Node^.Left) + CountNodes(Node^.Right);
end;

function TBinaryTree.Count: Integer;
begin
  Result := CountNodes(FRoot);
end;

function TBinaryTree.TreeHeight(Node: PTreeNode): Integer;
var
  LeftH, RightH: Integer;
begin
  if Node = nil then
    Result := 0
  else
  begin
    LeftH := TreeHeight(Node^.Left);
    RightH := TreeHeight(Node^.Right);
    if LeftH > RightH then
      Result := LeftH + 1
    else
      Result := RightH + 1;
  end;
end;

function TBinaryTree.Height: Integer;
begin
  Result := TreeHeight(FRoot);
end;

var
  Tree: TBinaryTree;
begin
  Tree := TBinaryTree.Create;
  Tree.Insert(50);
  Tree.Insert(30);
  Tree.Insert(70);
  Tree.Insert(20);
  Tree.Insert(40);
  Tree.Insert(60);
  Tree.Insert(80);

  Writeln('In-order:');
  Tree.PrintInOrder;
  Writeln('Count: ', Tree.Count);
  Writeln('Height: ', Tree.Height);
  Writeln('Contains 40? ', Tree.Contains(40));
  Writeln('Contains 100? ', Tree.Contains(100));

  Tree.Free;
end.

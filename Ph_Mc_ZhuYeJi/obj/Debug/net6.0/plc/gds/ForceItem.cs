// <auto-generated>
//     Generated by the protocol buffer compiler.  DO NOT EDIT!
//     source: Plc/Gds/ForceItem.proto
// </auto-generated>
#pragma warning disable 1591, 0612, 3021, 8981
#region Designer generated code

using pb = global::Google.Protobuf;
using pbc = global::Google.Protobuf.Collections;
using pbr = global::Google.Protobuf.Reflection;
using scg = global::System.Collections.Generic;
namespace Arp.Plc.Gds.Services.Grpc {

  /// <summary>Holder for reflection information generated from Plc/Gds/ForceItem.proto</summary>
  public static partial class ForceItemReflection {

    #region Descriptor
    /// <summary>File descriptor for Plc/Gds/ForceItem.proto</summary>
    public static pbr::FileDescriptor Descriptor {
      get { return descriptor; }
    }
    private static pbr::FileDescriptor descriptor;

    static ForceItemReflection() {
      byte[] descriptorData = global::System.Convert.FromBase64String(
          string.Concat(
            "ChdQbGMvR2RzL0ZvcmNlSXRlbS5wcm90bxIZQXJwLlBsYy5HZHMuU2Vydmlj",
            "ZXMuR3JwYxoOQXJwVHlwZXMucHJvdG8iUAoJRm9yY2VJdGVtEhQKDFZhcmlh",
            "YmxlTmFtZRgBIAEoCRItCgpGb3JjZVZhbHVlGAIgASgLMhkuQXJwLlR5cGUu",
            "R3JwYy5PYmplY3RUeXBlYgZwcm90bzM="));
      descriptor = pbr::FileDescriptor.FromGeneratedCode(descriptorData,
          new pbr::FileDescriptor[] { global::Arp.Type.Grpc.ArpTypesReflection.Descriptor, },
          new pbr::GeneratedClrTypeInfo(null, null, new pbr::GeneratedClrTypeInfo[] {
            new pbr::GeneratedClrTypeInfo(typeof(global::Arp.Plc.Gds.Services.Grpc.ForceItem), global::Arp.Plc.Gds.Services.Grpc.ForceItem.Parser, new[]{ "VariableName", "ForceValue" }, null, null, null, null)
          }));
    }
    #endregion

  }
  #region Messages
  /// <summary>
  //// &lt;summary>
  //// A force item structure.
  //// &lt;/summary>
  /// </summary>
  [global::System.Diagnostics.DebuggerDisplayAttribute("{ToString(),nq}")]
  public sealed partial class ForceItem : pb::IMessage<ForceItem>
  #if !GOOGLE_PROTOBUF_REFSTRUCT_COMPATIBILITY_MODE
      , pb::IBufferMessage
  #endif
  {
    private static readonly pb::MessageParser<ForceItem> _parser = new pb::MessageParser<ForceItem>(() => new ForceItem());
    private pb::UnknownFieldSet _unknownFields;
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public static pb::MessageParser<ForceItem> Parser { get { return _parser; } }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public static pbr::MessageDescriptor Descriptor {
      get { return global::Arp.Plc.Gds.Services.Grpc.ForceItemReflection.Descriptor.MessageTypes[0]; }
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    pbr::MessageDescriptor pb::IMessage.Descriptor {
      get { return Descriptor; }
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public ForceItem() {
      OnConstruction();
    }

    partial void OnConstruction();

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public ForceItem(ForceItem other) : this() {
      variableName_ = other.variableName_;
      forceValue_ = other.forceValue_ != null ? other.forceValue_.Clone() : null;
      _unknownFields = pb::UnknownFieldSet.Clone(other._unknownFields);
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public ForceItem Clone() {
      return new ForceItem(this);
    }

    /// <summary>Field number for the "VariableName" field.</summary>
    public const int VariableNameFieldNumber = 1;
    private string variableName_ = "";
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public string VariableName {
      get { return variableName_; }
      set {
        variableName_ = pb::ProtoPreconditions.CheckNotNull(value, "value");
      }
    }

    /// <summary>Field number for the "ForceValue" field.</summary>
    public const int ForceValueFieldNumber = 2;
    private global::Arp.Type.Grpc.ObjectType forceValue_;
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public global::Arp.Type.Grpc.ObjectType ForceValue {
      get { return forceValue_; }
      set {
        forceValue_ = value;
      }
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public override bool Equals(object other) {
      return Equals(other as ForceItem);
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public bool Equals(ForceItem other) {
      if (ReferenceEquals(other, null)) {
        return false;
      }
      if (ReferenceEquals(other, this)) {
        return true;
      }
      if (VariableName != other.VariableName) return false;
      if (!object.Equals(ForceValue, other.ForceValue)) return false;
      return Equals(_unknownFields, other._unknownFields);
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public override int GetHashCode() {
      int hash = 1;
      if (VariableName.Length != 0) hash ^= VariableName.GetHashCode();
      if (forceValue_ != null) hash ^= ForceValue.GetHashCode();
      if (_unknownFields != null) {
        hash ^= _unknownFields.GetHashCode();
      }
      return hash;
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public override string ToString() {
      return pb::JsonFormatter.ToDiagnosticString(this);
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public void WriteTo(pb::CodedOutputStream output) {
    #if !GOOGLE_PROTOBUF_REFSTRUCT_COMPATIBILITY_MODE
      output.WriteRawMessage(this);
    #else
      if (VariableName.Length != 0) {
        output.WriteRawTag(10);
        output.WriteString(VariableName);
      }
      if (forceValue_ != null) {
        output.WriteRawTag(18);
        output.WriteMessage(ForceValue);
      }
      if (_unknownFields != null) {
        _unknownFields.WriteTo(output);
      }
    #endif
    }

    #if !GOOGLE_PROTOBUF_REFSTRUCT_COMPATIBILITY_MODE
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    void pb::IBufferMessage.InternalWriteTo(ref pb::WriteContext output) {
      if (VariableName.Length != 0) {
        output.WriteRawTag(10);
        output.WriteString(VariableName);
      }
      if (forceValue_ != null) {
        output.WriteRawTag(18);
        output.WriteMessage(ForceValue);
      }
      if (_unknownFields != null) {
        _unknownFields.WriteTo(ref output);
      }
    }
    #endif

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public int CalculateSize() {
      int size = 0;
      if (VariableName.Length != 0) {
        size += 1 + pb::CodedOutputStream.ComputeStringSize(VariableName);
      }
      if (forceValue_ != null) {
        size += 1 + pb::CodedOutputStream.ComputeMessageSize(ForceValue);
      }
      if (_unknownFields != null) {
        size += _unknownFields.CalculateSize();
      }
      return size;
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public void MergeFrom(ForceItem other) {
      if (other == null) {
        return;
      }
      if (other.VariableName.Length != 0) {
        VariableName = other.VariableName;
      }
      if (other.forceValue_ != null) {
        if (forceValue_ == null) {
          ForceValue = new global::Arp.Type.Grpc.ObjectType();
        }
        ForceValue.MergeFrom(other.ForceValue);
      }
      _unknownFields = pb::UnknownFieldSet.MergeFrom(_unknownFields, other._unknownFields);
    }

    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    public void MergeFrom(pb::CodedInputStream input) {
    #if !GOOGLE_PROTOBUF_REFSTRUCT_COMPATIBILITY_MODE
      input.ReadRawMessage(this);
    #else
      uint tag;
      while ((tag = input.ReadTag()) != 0) {
        switch(tag) {
          default:
            _unknownFields = pb::UnknownFieldSet.MergeFieldFrom(_unknownFields, input);
            break;
          case 10: {
            VariableName = input.ReadString();
            break;
          }
          case 18: {
            if (forceValue_ == null) {
              ForceValue = new global::Arp.Type.Grpc.ObjectType();
            }
            input.ReadMessage(ForceValue);
            break;
          }
        }
      }
    #endif
    }

    #if !GOOGLE_PROTOBUF_REFSTRUCT_COMPATIBILITY_MODE
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute]
    [global::System.CodeDom.Compiler.GeneratedCode("protoc", null)]
    void pb::IBufferMessage.InternalMergeFrom(ref pb::ParseContext input) {
      uint tag;
      while ((tag = input.ReadTag()) != 0) {
        switch(tag) {
          default:
            _unknownFields = pb::UnknownFieldSet.MergeFieldFrom(_unknownFields, ref input);
            break;
          case 10: {
            VariableName = input.ReadString();
            break;
          }
          case 18: {
            if (forceValue_ == null) {
              ForceValue = new global::Arp.Type.Grpc.ObjectType();
            }
            input.ReadMessage(ForceValue);
            break;
          }
        }
      }
    }
    #endif

  }

  #endregion

}

#endregion Designer generated code
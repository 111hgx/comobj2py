import win32com.client
import pythoncom
from typing import List, Dict, Any
import os
import sys
import win32com.client.makepy
import traceback

def get_com_interfaces(com_object_name: str) -> List[str]:
    """获取COM对象的所有属性和方法名称"""
    try:
        pythoncom.CoInitialize()
        print(f"\n正在创建COM对象: {com_object_name}")
        
        # 生成类型库
        try:
            print("正在生成类型库...")
            tlb = win32com.client.makepy.GenerateFromTypeLibSpec(com_object_name)
            print(f"已生成类型库: {tlb}")
        except Exception as e:
            print(f"生成类型库时出错: {e}")
            print("继续尝试直接创建COM对象...")
        
        # 创建COM对象
        print("正在创建COM对象...")
        com_object = win32com.client.Dispatch(com_object_name)
        print("COM对象创建成功")
        
        print("\nCOM对象的所有属性和方法:")
        all_attrs = dir(com_object)
        interfaces = []
        
        # 获取类型信息
        try:
            print("正在获取类型信息...")
            type_info = com_object._oleobj_.GetTypeInfo()
            if type_info:
                type_attr = type_info.GetTypeAttr()
                print(f"\n类型信息:")
                print(f"- 类型名称: {type_attr.tdescAlias}")
                print(f"- 函数数量: {type_attr.cFuncs}")
                print(f"- 变量数量: {type_attr.cVars}")
        except Exception as e:
            print(f"获取类型信息时出错: {e}")
        
        print("\n正在分析属性和方法...")
        for attr in all_attrs:
            if not attr.startswith('_'):
                try:
                    value = getattr(com_object, attr)
                    print(f"- {attr}: {type(value)}")
                    interfaces.append(attr)
                except Exception as e:
                    print(f"- {attr}: Error - {str(e)}")
        
        return interfaces
    except Exception as e:
        print(f"Error: {e}")
        print("详细错误信息:")
        traceback.print_exc()
        return []
    finally:
        pythoncom.CoUninitialize()

def get_interface_members(com_object_name: str, interface_name: str) -> Dict[str, Dict[str, Any]]:
    """获取指定接口的所有属性和方法"""
    try:
        pythoncom.CoInitialize()
        print(f"\n正在获取接口 {interface_name} 的成员...")
        com_object = win32com.client.Dispatch(com_object_name)
        interface = getattr(com_object, interface_name)
        
        members = {}
        
        # 获取类型信息
        try:
            print("正在获取接口类型信息...")
            type_info = interface._oleobj_.GetTypeInfo()
            if type_info:
                type_attr = type_info.GetTypeAttr()
                print(f"\n接口类型信息:")
                print(f"- 函数数量: {type_attr.cFuncs}")
                print(f"- 变量数量: {type_attr.cVars}")
        except Exception as e:
            print(f"获取接口类型信息时出错: {e}")
        
        print("\n正在分析接口成员...")
        for member_name in dir(interface):
            if not member_name.startswith('_'):
                try:
                    member = getattr(interface, member_name)
                    if callable(member):
                        doc = getattr(member, '__doc__', '')
                        if doc:
                            doc = doc.replace('\\', '\\\\')
                        members[member_name] = {
                            'name': member_name,
                            'type': 'method',
                            'doc': doc
                        }
                        print(f"- 方法: {member_name}")
                    else:
                        try:
                            value = getattr(interface, member_name)
                            value_str = str(value).replace('\\', '\\\\')
                            members[member_name] = {
                                'name': member_name,
                                'type': 'property',
                                'value': value_str
                            }
                            print(f"- 属性: {member_name} = {value_str}")
                        except:
                            members[member_name] = {
                                'name': member_name,
                                'type': 'property',
                                'value': 'Unknown'
                            }
                            print(f"- 属性: {member_name} = Unknown")
                except Exception as e:
                    print(f"处理成员 {member_name} 时出错: {str(e)}")
                    continue
        return members
    except Exception as e:
        print(f"Error: {e}")
        print("详细错误信息:")
        traceback.print_exc()
        return {}
    finally:
        pythoncom.CoUninitialize()

def generate_interface_code(com_object_name: str, interface_name: str) -> str:
    """生成接口的Python代码"""
    members = get_interface_members(com_object_name, interface_name)
    
    code = f'''# -*- coding: utf-8 -*-
import win32com.client
import pythoncom
from typing import Any, Optional

class {interface_name}:
    """
    Wrapper class for {interface_name} interface of {com_object_name}
    """
    def __init__(self, com_object_name: str = "{com_object_name}"):
        pythoncom.CoInitialize()
        self.com_object = win32com.client.Dispatch(com_object_name)
        self.interface = getattr(self.com_object, "{interface_name}")
    
    def __del__(self):
        pythoncom.CoUninitialize()
    
    def _get_member(self, name: str) -> Any:
        """获取成员值"""
        return getattr(self.interface, name)
    
    def _set_member(self, name: str, value: Any) -> None:
        """设置成员值"""
        setattr(self.interface, name, value)
    
'''
    
    # 添加属性
    for member_name, member_info in members.items():
        if member_info['type'] == 'property':
            code += f'''    @property
    def {member_name}(self) -> Any:
        """
        Property: {member_name}
        Current value: {member_info['value']}
        """
        return self._get_member("{member_name}")
    
    @{member_name}.setter
    def {member_name}(self, value: Any) -> None:
        """设置 {member_name} 属性值"""
        self._set_member("{member_name}", value)
    
'''
    
    # 添加方法
    for member_name, member_info in members.items():
        if member_info['type'] == 'method':
            doc = member_info['doc'] or f"Call {member_name} method"
            code += f'''    def {member_name}(self, *args, **kwargs) -> Any:
        """
        {doc}
        """
        return self.interface.{member_name}(*args, **kwargs)
    
'''
    
    return code

def generate_all_interfaces(com_object_name: str, output_dir: str = "generated"):
    """生成所有接口的代码"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    print(f"\n正在获取 {com_object_name} 的属性和方法列表...")
    interfaces = get_com_interfaces(com_object_name)
    
    if not interfaces:
        print("未找到任何属性和方法")
        return
    
    print(f"\n找到 {len(interfaces)} 个属性和方法:")
    for interface in interfaces:
        print(f"- {interface}")
    
    print("\n开始生成代码...")
    
    # 生成__init__.py
    init_code = "# -*- coding: utf-8 -*-\nfrom typing import List\n\n"
    
    for interface_name in interfaces:
        print(f"\n正在生成 {interface_name} 接口代码...")
        code = generate_interface_code(com_object_name, interface_name)
        
        # 保存到文件
        filename = f"{interface_name.lower()}.py"
        filepath = os.path.join(output_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(code)
        print(f"已保存到 {filepath}")
        
        # 添加到__init__.py
        init_code += f"from .{interface_name.lower()} import {interface_name}\n"
    
    # 保存__init__.py
    init_path = os.path.join(output_dir, "__init__.py")
    with open(init_path, 'w', encoding='utf-8') as f:
        f.write(init_code)
    print(f"\n已生成 __init__.py")
    
    # 显示使用示例
    print("\n使用示例:")
    print(f"""
from {output_dir} import {interfaces[0]}

# 创建接口实例
interface = {interfaces[0]}()

# 访问属性
value = interface.SomeProperty

# 调用方法
result = interface.SomeMethod(param1, param2)
    """)

def main():
    com_object_name = input("请输入COM对象名称: ")
    generate_all_interfaces(com_object_name)

if __name__ == "__main__":
    main() 
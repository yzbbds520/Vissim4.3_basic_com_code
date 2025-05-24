# Python脚本用于探索Vissim 4.3/4.5中创建Link的方法
import win32com.client as com
import traceback


# 辅助函数：打印对象的方法和属性
def explore_object(obj, obj_name, depth=0, max_depth=2):
    if depth > max_depth:
        return

    indent = "  " * depth
    print(f"\n{indent}{'=' * 40}")
    print(f"{indent}探索对象: {obj_name}")
    print(f"{indent}{'=' * 40}")

    # 尝试获取对象的所有属性和方法
    try:
        attrs = dir(obj)
        print(f"\n{indent}对象 {obj_name} 的属性和方法:")
        for attr in sorted(attrs):
            if not attr.startswith('_'):  # 忽略内部属性
                print(f"{indent}  - {attr}")

                # 尝试获取属性值（只对可能的集合或子对象进行探索）
                if attr in ['Links', 'Net', 'AddLink', 'CreateLink', 'AddLinks', 'Points']:
                    try:
                        attr_value = getattr(obj, attr)
                        if depth < max_depth:  # 限制递归深度
                            if callable(attr_value):
                                print(f"{indent}    (这是一个方法)")
                            else:
                                print(f"{indent}    (这是一个属性/对象)")
                                explore_object(attr_value, f"{obj_name}.{attr}", depth + 1, max_depth)
                    except:
                        print(f"{indent}    (无法获取值)")
    except:
        print(f"{indent}无法获取 {obj_name} 的属性和方法列表")

    # 尝试获取对象的类型信息
    try:
        type_info = str(type(obj))
        print(f"\n{indent}对象类型: {type_info}")
    except:
        print(f"{indent}无法获取对象类型信息")


try:
    # 连接到VISSIM
    print("正在连接到VISSIM 4.3/4.5...")
    Vissim = com.Dispatch("VISSIM.Vissim")
    print("成功连接到VISSIM")

    # 加载网络文件
    try:
        Filename = r"D:\python\project\vissim\try.inp"
        print(f"尝试加载网络文件: {Filename}")
        Vissim.LoadNet(Filename)
        print("成功加载网络文件")
    except Exception as e:
        print(f"加载网络文件失败: {str(e)}")
        print("尝试创建新网络...")
        Vissim.New()
        print("已创建新网络")

    # 探索Net对象
    print("\n开始探索Net对象...")
    explore_object(Vissim.Net, "Vissim.Net")

    # 特别检查Links集合
    try:
        if hasattr(Vissim.Net, 'Links'):
            print("\n\n特别探索Links集合...")
            explore_object(Vissim.Net.Links, "Vissim.Net.Links", 0, 1)

            # 尝试获取Links的数量
            try:
                links_count = Vissim.Net.Links.Count
                print(f"\n当前Links数量: {links_count}")
            except:
                print("\n无法获取Links数量")
    except Exception as e:
        print(f"\n探索Links集合时出错: {str(e)}")

    # 尝试不同方式创建Link
    print("\n\n尝试不同方式创建Link:")

    # 方法1: 尝试通过AddLink方法
    try:
        print("\n方法1: 尝试Net.Links.AddLink")
        if hasattr(Vissim.Net.Links, 'AddLink'):
            print("找到Net.Links.AddLink方法，尝试调用...")
            # 尝试不同的参数组合
            try:
                # 创建从(0,0)到(0,50)的链接
                new_link = Vissim.Net.Links.AddLink(0, 0, 0, 50, 1)
                print(f"成功创建链接！链接ID: {new_link.ID if hasattr(new_link, 'ID') else '未知'}")
            except Exception as e:
                print(f"调用Net.Links.AddLink(0, 0, 0, 50, 1)失败: {str(e)}")
                print("尝试不同的参数...")
                try:
                    new_link = Vissim.Net.Links.AddLink("0 0 0 50")
                    print(f"成功创建链接！链接ID: {new_link.ID if hasattr(new_link, 'ID') else '未知'}")
                except Exception as e:
                    print(f"调用Net.Links.AddLink(\"0 0 0 50\")失败: {str(e)}")
        else:
            print("Net.Links对象没有AddLink方法")
    except Exception as e:
        print(f"方法1测试失败: {str(e)}")
        print(traceback.format_exc())

    # 方法2: 尝试通过CreateLink方法
    try:
        print("\n方法2: 尝试Net.CreateLink")
        if hasattr(Vissim.Net, 'CreateLink'):
            print("找到Net.CreateLink方法，尝试调用...")
            try:
                new_link = Vissim.Net.CreateLink(0, 0, 0, 50, 1)
                print(f"成功创建链接！链接ID: {new_link.ID if hasattr(new_link, 'ID') else '未知'}")
            except Exception as e:
                print(f"调用Net.CreateLink(0, 0, 0, 50, 1)失败: {str(e)}")
        else:
            print("Net对象没有CreateLink方法")
    except Exception as e:
        print(f"方法2测试失败: {str(e)}")

    # 方法3: 检查是否有Points或Coordinates方法
    try:
        print("\n方法3: 探索是否有点或坐标相关方法")
        # 检查是否有Points集合
        if hasattr(Vissim.Net, 'Points'):
            print("找到Points集合，进一步探索...")
            explore_object(Vissim.Net.Points, "Vissim.Net.Points", 0, 1)
    except Exception as e:
        print(f"方法3测试失败: {str(e)}")

    # 探索可能的其他相关方法
    try:
        print("\n寻找其他可能的创建Link的方法...")
        net_attrs = dir(Vissim.Net)
        for attr in net_attrs:
            if not attr.startswith('_') and (
                    'link' in attr.lower() or 'create' in attr.lower() or 'add' in attr.lower()):
                print(f"  - 可能相关的方法/属性: {attr}")
    except Exception as e:
        print(f"探索其他方法时出错: {str(e)}")

except Exception as e:
    print(f"发生错误: {str(e)}")
    print(traceback.format_exc())
finally:
    try:
        # 尝试关闭VISSIM
        print("\n尝试关闭VISSIM...")
        Vissim = None
        print("VISSIM已关闭")
    except:
        pass
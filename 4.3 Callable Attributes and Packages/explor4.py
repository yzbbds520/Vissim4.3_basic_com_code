# Python脚本用于探索VISSIM 4.3/4.5中的跟驰模型相关接口
import win32com.client as com
import traceback


# 辅助函数：探索对象的属性和方法
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

                # 尝试获取属性值（只对可能的跟驰模型相关对象进行探索）
                keywords = ['follow', 'driver', 'behavior', 'model', 'wiedemann', 'car', 'vehicle']
                if any(keyword in attr.lower() for keyword in keywords) and depth < max_depth:
                    try:
                        attr_value = getattr(obj, attr)
                        if callable(attr_value):
                            print(f"{indent}    (这是一个方法)")
                        else:
                            print(f"{indent}    (这是一个属性/对象)")
                            explore_object(attr_value, f"{obj_name}.{attr}", depth + 1, max_depth)
                    except Exception as e:
                        print(f"{indent}    (无法获取值: {str(e)})")
    except Exception as e:
        print(f"{indent}无法获取 {obj_name} 的属性和方法列表: {str(e)}")


# 辅助函数：探索对象的AttValue可能的属性
def explore_att_values(obj, obj_name):
    print(f"\n{'=' * 50}")
    print(f"探索 {obj_name} 可能的AttValue属性:")
    print(f"{'=' * 50}")

    # 常见的跟驰模型相关参数
    car_following_params = [
        # Wiedemann 74模型参数
        'W74ax', 'W74bxAdd', 'W74bxMult', 'W74cc',
        # Wiedemann 99模型参数
        'W99cc0', 'W99cc1', 'W99cc2', 'W99cc3', 'W99cc4', 'W99cc5',
        'W99cc6', 'W99cc7', 'W99cc8', 'W99cc9',
        # 其他常见参数（使用不同名称）
        'CC0', 'CC1', 'CC2', 'CC3', 'CC4', 'CC5', 'CC6', 'CC7', 'CC8', 'CC9',
        'AX', 'BX', 'CX', 'SafetyDist', 'FollowDist', 'StandStill', 'Headway',
        'DriverModel', 'CarFollow', 'LookAheadDist', 'LookBackDist',
        'ModelType', 'FollowingModel', 'DriverBehavior', 'FollowingBehavior'
    ]

    if hasattr(obj, 'AttValue'):
        for param in car_following_params:
            try:
                value = obj.AttValue(param)
                print(f"  - {param}: {value}")
            except:
                # 参数不存在，继续尝试下一个
                pass


try:
    # 连接到VISSIM
    print("正在连接到VISSIM 4.3/4.5...")
    Vissim = com.Dispatch("VISSIM.Vissim")
    print("成功连接到VISSIM")

    # 加载网络文件或创建新网络
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

    # 探索VISSIM主对象
    print("\n开始探索VISSIM对象...")
    explore_object(Vissim, "Vissim", max_depth=1)

    # 特别探索Net对象
    print("\n深入探索Net对象...")
    if hasattr(Vissim, 'Net'):
        explore_object(Vissim.Net, "Vissim.Net", max_depth=1)

    # 探索与驾驶行为相关的对象
    driving_behavior_objects = [
        ('DrivingBehaviorParSets', 'Vissim.Net.DrivingBehaviorParSets'),
        ('DrivingBehaviors', 'Vissim.Net.DrivingBehaviors'),
        ('DriverBehavior', 'Vissim.Net.DriverBehavior'),
        ('DriverModels', 'Vissim.Net.DriverModels'),
        ('DriverBehaviorModels', 'Vissim.Net.DriverBehaviorModels')
    ]

    print("\n探索与驾驶行为相关的对象...")
    for attr_name, full_path in driving_behavior_objects:
        try:
            if hasattr(Vissim.Net, attr_name):
                obj = getattr(Vissim.Net, attr_name)
                explore_object(obj, full_path, max_depth=2)

                # 尝试获取第一个驾驶行为对象
                try:
                    if hasattr(obj, 'Count') and obj.Count > 0:
                        if hasattr(obj, 'Item'):
                            first_item = obj.Item(1)
                            print(f"\n尝试获取第一个{attr_name}项...")
                            explore_object(first_item, f"{full_path}.Item(1)", max_depth=1)
                            explore_att_values(first_item, f"{full_path}.Item(1)")
                        elif hasattr(obj, 'ItemByKey'):
                            first_item = obj.ItemByKey(1)
                            print(f"\n尝试获取第一个{attr_name}项...")
                            explore_object(first_item, f"{full_path}.ItemByKey(1)", max_depth=1)
                            explore_att_values(first_item, f"{full_path}.ItemByKey(1)")
                except Exception as e:
                    print(f"获取{attr_name}第一项失败: {str(e)}")
        except Exception as e:
            print(f"探索{attr_name}失败: {str(e)}")

    # 探索车辆对象，看是否有跟驰相关属性
    print("\n探索车辆对象的跟驰相关属性...")
    try:
        vehicles = Vissim.Net.Vehicles
        if hasattr(vehicles, 'Count') and vehicles.Count > 0:
            if hasattr(vehicles, 'Item'):
                first_veh = vehicles.Item(1)
                print("\n尝试获取第一个车辆...")
                explore_object(first_veh, "Vissim.Net.Vehicles.Item(1)", max_depth=1)
                explore_att_values(first_veh, "Vissim.Net.Vehicles.Item(1)")
            elif hasattr(vehicles, 'GetAll'):
                all_vehs = vehicles.GetAll()
                if len(all_vehs) > 0:
                    first_veh = all_vehs[0]
                    print("\n尝试获取第一个车辆...")
                    explore_object(first_veh, "Vissim.Net.Vehicles.GetAll()[0]", max_depth=1)
                    explore_att_values(first_veh, "Vissim.Net.Vehicles.GetAll()[0]")
    except Exception as e:
        print(f"探索车辆对象失败: {str(e)}")

    # 检查是否有任何API文档或辅助方法
    print("\n检查是否有文档相关的方法...")
    api_methods = ['Help', 'Documentation', 'COM', 'COMHelp', 'APIHelp', 'GetAPIReference']
    for method in api_methods:
        if hasattr(Vissim, method):
            print(f"  - 找到可能的API文档方法: Vissim.{method}")

    # 尝试查找任何可能的自定义模型接口
    print("\n查找可能的自定义模型接口...")
    custom_interfaces = ['CustomModels', 'UserModels', 'ExternalDriverModel', 'DriverModelInterface', 'API']
    for interface in custom_interfaces:
        if hasattr(Vissim, interface):
            print(f"  - 找到可能的自定义接口: Vissim.{interface}")
            obj = getattr(Vissim, interface)
            explore_object(obj, f"Vissim.{interface}", max_depth=1)
        elif hasattr(Vissim.Net, interface):
            print(f"  - 找到可能的自定义接口: Vissim.Net.{interface}")
            obj = getattr(Vissim.Net, interface)
            explore_object(obj, f"Vissim.Net.{interface}", max_depth=1)

    # 搜索所有与car following相关的属性或方法
    print("\n搜索所有与car following相关的属性和方法...")
    keywords = ['follow', 'driver', 'behavior', 'model', 'wiedemann']

    # 在Vissim对象中搜索
    vissim_attrs = dir(Vissim)
    for attr in vissim_attrs:
        if not attr.startswith('_') and any(keyword in attr.lower() for keyword in keywords):
            print(f"  - 在Vissim中找到可能相关的属性/方法: {attr}")

    # 在Vissim.Net对象中搜索
    if hasattr(Vissim, 'Net'):
        net_attrs = dir(Vissim.Net)
        for attr in net_attrs:
            if not attr.startswith('_') and any(keyword in attr.lower() for keyword in keywords):
                print(f"  - 在Vissim.Net中找到可能相关的属性/方法: {attr}")

except Exception as e:
    print(f"发生错误: {str(e)}")
    print(traceback.format_exc())
finally:
    try:
        # 关闭VISSIM
        print("\n尝试关闭VISSIM...")
        Vissim = None
        print("VISSIM已关闭")
    except:
        print("关闭VISSIM时出错")
from winreg import*
import platform
from datetime import*


def out_info(info_pc):
    # Выводим информацию в файл из info_all_pc
    if len(info_pc) != 0:
        with open('info_pc.txt', 'w') as inf:
            for key in info_pc.keys():
                counter = 0
                for key1, value1 in info_pc[key].items():
                    if counter == 4 or counter == 7:
                        inf.write('\n')
                    inf.write(key1 + value1 + '\n')
                    counter += 1
                inf.write('\n' + '*' * 20 + '\n' * 2)


def out_info_error(error_conect_pc):
    # Выводим информацию в файл из bad_pc
    if len(error_conect_pc) != 0:
        with open('error_conect_pc.txt', 'w') as inf:
            for i in error_conect_pc:
                inf.write(str(i) + '\n')


def creat_list_pc():
    # create name pc
    list_pc = []
    with open('name_pc.txt', encoding='utf-8') as pc:
        for line in pc:
            list_pc.append(line.strip())
    return list_pc


def info_work_stash(aReg, other_view_flag):
    try:
        # Info about PC
        aKey = OpenKeyEx(aReg, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion", access=KEY_READ | other_view_flag)
        Oc = QueryValueEx(aKey, 'ProductName')
        CloseKey(aKey)

        aKey = OpenKeyEx(aReg, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", access=KEY_READ | other_view_flag)
        Arh_Platf = QueryValueEx(aKey, 'PROCESSOR_ARCHITECTURE')
        CloseKey(aKey)

        aKey = OpenKeyEx(aReg, r"SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", access=KEY_READ | other_view_flag)
        NamePC = QueryValueEx(aKey, 'ComputerName')
        CloseKey(aKey)
        return NamePC, Oc, Arh_Platf
    except:
        return ['No'], ['No'], [' ']


def info_criptopro(aReg, other_view_flag):
    try:
        # Info about CriptoPro
        aKey = OpenKeyEx(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\7AB5E7046046FB044ACD63458B5F481C\InstallProperties", access=KEY_READ | other_view_flag)
        ProductID = QueryValueEx(aKey, 'ProductID')
        DisplayVersion = QueryValueEx(aKey, 'DisplayVersion')
        InstallDate = QueryValueEx(aKey, 'InstallDate')
        CloseKey(aKey)
        inst = InstallDate[0]
        InstallDate = inst[6 : 8] + '.' + inst[4 : 6] + '.' + inst[0 : 4]
        return ProductID, DisplayVersion, InstallDate
    except:
        return ['No'], ['No'], 'No'


def info_hd(aReg, other_view_flag):
    try:
        # Info Hardriver
        aKey = OpenKeyEx(aReg, r"HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 1\Logical Unit Id 0", access=KEY_READ | KEY_WOW64_32KEY)
        Identifier = QueryValueEx(aKey, 'Identifier')
        SerialNumber = QueryValueEx(aKey, 'SerialNumber')
        CloseKey(aKey)
        return Identifier, SerialNumber
    except:
        return ['No'], ['No']


def sbor_sved(name_pc):
    # разрядность системы PC на котором запускается программа
    bitness = platform.architecture()[0]
    if bitness == '32bit':
        other_view_flag = KEY_WOW64_64KEY
    elif bitness == '64bit':
        other_view_flag = KEY_WOW64_32KEY

    # текущая дата
    dattim = date.today()
    name = '\\\\' + name_pc
    print(name)
    # r"{}".format(name)
    try:
        # проверка подключения к удаленному компьютеру
        aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)

        NamePC, Oc, Arh_Platf = info_work_stash(aReg, other_view_flag)

        ProductID, DisplayVersion, InstallDate = info_criptopro(aReg, other_view_flag)

        Identifier, SerialNumber = info_hd(aReg, other_view_flag)

        info_pc = {'Name PC:    ': NamePC[0], 'Oc:         ': Oc[0] + ' '
                   + Arh_Platf[0], 'Data Check: ': str(dattim),
                   'Time Check: ': str(datetime.now().time())[: 8],
                   'ProductID:  ': ProductID[0],
                   'DispVeion:  ': DisplayVersion[0],
                   'InstDate:   ': InstallDate,
                   'Hardriver   ': Identifier[0],
                   'S/N         ': SerialNumber[0].strip()
                   }
        return info_pc
    except:
        return name_pc


# список компютеров к которым не смогли подключиться
error_conect_pc = []
info_pc = {}
# создает список имен компьютеров
list_name_pc = creat_list_pc()
# Проверка типа list or dict
for name in list_name_pc:
    result = sbor_sved(name)
    if type(result) == dict:
        info_pc[name] = result
    else:
        error_conect_pc.append(result)

out_info(info_pc)

out_info_error(error_conect_pc)

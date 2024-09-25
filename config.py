import configparser


class ConfigHandler:
    def __init__(self, config_file='config.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.load_config()

    def load_config(self):
        self.config.read(self.config_file)

    def get_value(self, section, key):
        if section in self.config and key in self.config[section]:
            return self.config[section][key]
        else:
            return None

    def set_value(self, section, key, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, value)
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)


if __name__ == '__main__':
    handler = ConfigHandler()

    # 获取配置值
    value = handler.get_value('new_section', 'new_key')
    print(value)

    # 设置配置值
    handler.set_value('new_section', 'new_key', 'new_value')
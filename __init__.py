from powercfg import PowerCfg

if __name__ == '__main__':
    cfg = PowerCfg()
    cfg.load_from_json('schema.json')
    cfg.apply_schema()
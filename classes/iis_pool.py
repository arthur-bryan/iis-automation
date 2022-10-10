class IISPool:

    def __init__(self, hostname, name, state, auto_start, start_mode):
        self.hostname = hostname
        self.name = name
        self.state = state
        self.auto_start = auto_start
        self.start_mode = start_mode

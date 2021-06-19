
class DeviceModelDoesnotExistException(Exception):

    def __str__(self):
        return "Target device model doesn't exist."


class ParameterCannotBeNone(Exception):

    def __str__(self):
        return "Parameter cannot all be None."


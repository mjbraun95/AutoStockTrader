from ib.opt import ibConnection, message

def my_account_handler(msg):
    ... do something with account msg ...

def my_tick_handler(msg):
    ... do something with market data msg ...

connection = ibConnection()
connection.register(my_account_handler, 'UpdateAccountValue')
connection.register(my_tick_handler, 'TickSize', 'TickPrice')
connection.connect()
connection.reqAccountUpdates(...)

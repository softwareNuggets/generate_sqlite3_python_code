import sqlite3
import csv

def add_player_to_database(conn, cur):

    cur.execute("""
        CREATE TABLE if not exists Player
        (
            player_id       integer,
            player_name     text,
            position_id     integer,
            annual_salary       decimal(10,2),
            PRIMARY KEY (player_id)
        )
        """)
    conn.commit()


def load_Player(conn, cur, filename):

    with open(filename,"r") as fh:

        reader = csv.reader(fh, delimiter=',')

        # if the file DOES NOT have a header row use this line
        next(csv.reader(fh), None)  # skip first row


        stmt = "insert into Player("
        stmt += "player_id,player_name,position_id,annual_salary)"
        stmt += " values (?,?,?,?)"

        for entry in reader:

            try:

                record = (entry[0],entry[1],entry[2],entry[3])

                cur.execute(stmt,record)

                conn.commit()

            except Exception as err:

                print(f'Line:{reader.line_num}, Record: {record}')
                print(f'Exception: {err}')



if __name__ == "__main__":

    conn     = sqlite3.connect("temp.db")
    cur      = conn.cursor()

    add_table_to_database(conn, cur)
    load_Player(conn, cur, "C:\\Youtube\\python\\players\\team.csv")
    conn.close()



import sqlite3
import csv

def setup():
    # Create & connect to database
    db = sqlite3.connect("emails.db")

    # Create tables for email parser to fill
    db.execute("""
    CREATE TABLE IF NOT EXISTS "articles" (
	"id"	INTEGER,
	"date"	TEXT,
	"title"	TEXT UNIQUE,
	"publication"	TEXT,
	"tier"	INTEGER,
	"category"	TEXT,
	PRIMARY KEY("id" AUTOINCREMENT))
    """)

    db.execute("""
    CREATE TABLE IF NOT EXISTS "links" (
	"article_id"	INTEGER,
	"link0"	TEXT,
	"link1"	TEXT,
	"link2"	TEXT,
	"link3"	TEXT,
	"link4"	TEXT,
	PRIMARY KEY("article_id"))
    """)

    db.execute("""
    CREATE TABLE IF NOT EXISTS "medialist" (
	"url"	TEXT UNIQUE,
	"tier"	INTEGER,
	"type"	TEXT)
    """)

    db.execute("""
    CREATE TABLE IF NOT EXISTS "platforms" (
	"article_id"	INTEGER,
	"platform0"	TEXT,
	"platform1"	TEXT,
	"platform2"	TEXT,
	"platform3"	TEXT,
	"platform4"	TEXT,
	PRIMARY KEY("article_id"))
    """)

    # Write files
    results = db.execute("""SELECT * FROM medialist """)
    if not [r for r in results]:
        print("Error: 'medialist' table is empty!")
        print("Attempting to write medialist.csv to database")

        # Open new CSV file for writing
        try:
            with open("medialist.csv", "r") as medialist:
                reader = csv.DictReader(medialist)
                for row in reader:
                    print(row)
                    db.execute("INSERT INTO medialist (url, tier, type) VALUES(?, ?, ?)",
                                (row["url"], row["tier"], row["type"]))

        except FileNotFoundError:
            print("medialist.csv not found")
            exit(1)

    db.commit()

    # Reset databases after each run
    db.execute("DELETE FROM articles")
    db.execute("DELETE FROM platforms")
    db.execute("DELETE FROM links")

    return db

setup()
CREATE TABLE [T_Info] (
  [ID_Info] VARCHAR (15) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [InfoTitre] VARCHAR (100),
  [InfoTexte] LONGTEXT ,
  [ID_Lang] LONG ,
  [Code] LONGTEXT 
)

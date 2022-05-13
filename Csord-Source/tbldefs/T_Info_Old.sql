CREATE TABLE [T_Info_Old] (
  [ID_Info] VARCHAR (15) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [InfoTitre] VARCHAR (100),
  [InfoTexte] LONGTEXT ,
  [ID_Lang] LONG ,
  [pjCode] VARCHAR ,
  [Code] LONGTEXT 
)

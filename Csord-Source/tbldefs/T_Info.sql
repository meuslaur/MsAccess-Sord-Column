﻿CREATE TABLE [T_Info] (
  [ID_Info] VARCHAR (15) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ID_Lang] LONG ,
  [InfoTitre] VARCHAR (255),
  [InfoTexte] LONGTEXT ,
  [Code] LONGTEXT ,
  [ID_Res] SHORT 
)

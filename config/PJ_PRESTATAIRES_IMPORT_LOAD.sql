USE [aeosdb]
GO
/****** Object:  StoredProcedure [dbo].[PJ_PRESTATAIRES_IMPORT_LOAD]    Script Date: 02.02.2026 12:32:27 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER   PROCEDURE [dbo].[PJ_PRESTATAIRES_IMPORT_LOAD]
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @RunId UNIQUEIDENTIFIER = NEWID();
    DECLARE @RunTs DATETIME2(0) = SYSDATETIME();

    DECLARE @RowsInStage INT = (SELECT COUNT(*) FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE);
    DECLARE @PersonsQueued INT = 0;
    DECLARE @BlocksQueued INT = 0;

    DECLARE @SourceFile NVARCHAR(255) =
        (SELECT TOP (1) source_file FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE ORDER BY load_ts DESC);
    DECLARE @BatchId UNIQUEIDENTIFIER =
        (SELECT TOP (1) batch_id FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE ORDER BY load_ts DESC);

    BEGIN TRY
        BEGIN TRAN;

        /* 1) ENVOI DES PERSONNES (FUNCTION 9) */
        INSERT INTO dbo.[import] (
            import_function,      -- 9 = personnes
            carriertype,          -- 5 = prestataires
            lastname,
            initials,
            personnelnr,
            company,              -- libelle fournisseur
            vendor_code,          -- code fournisseur
            arrivaldatetime,
            leavedatetime,
            validfrom,
            validto,
            disabled
        )
        SELECT
            '9' AS import_function,
            '5' AS carriertype,
            s.lastname,
            NULLIF(LTRIM(RTRIM(COALESCE(s.firstname, s.initials))), '') AS initials,
            s.personnelnr,
            NULLIF(LTRIM(RTRIM(COALESCE(s.company, s.vendor_code))), '') AS company,
            NULLIF(LTRIM(RTRIM(s.vendor_code)), '') AS vendor_code,
            s.arrivaldate AS arrivaldatetime,
            s.leavedate   AS leavedatetime,
            NULL AS validfrom,
            NULL AS validto,
            0 AS disabled
        FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE AS s
        WHERE s.personnelnr IS NOT NULL
          AND s.lastname IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.vendor_code)), '') IS NOT NULL
          AND s.arrivaldate IS NOT NULL
          AND s.leavedate IS NOT NULL;

        SET @PersonsQueued = @@ROWCOUNT;

        /* 2) CALCUL DES PERSONNES A BLOQUER (FUNCTION 6) */
        IF OBJECT_ID('tempdb..#ToBlock','U') IS NOT NULL
            DROP TABLE #ToBlock;

        ;WITH CurrentFile AS (
            SELECT DISTINCT
                s.personnelnr AS Matricule,
                NULLIF(LTRIM(RTRIM(s.vendor_code)), '') AS VendorCode
            FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE AS s
            WHERE s.personnelnr IS NOT NULL
              AND NULLIF(LTRIM(RTRIM(s.vendor_code)), '') IS NOT NULL
        )
        SELECT
            t.Matricule,
            t.VendorCode
        INTO #ToBlock
        FROM dbo.PJ_PRESTATAIRES_IMPORT_TAMPON AS t
        LEFT JOIN CurrentFile AS c
               ON c.Matricule  = t.Matricule
              AND c.VendorCode = t.VendorCode
        WHERE c.Matricule IS NULL;

        INSERT INTO dbo.[import] (
            import_function,   -- 6 = blocage
            carriertype,       -- 5 = prestataires
            personnelnr,
            vendor_code,
            blocked,
            disabled
        )
        SELECT
            6 AS import_function,
            5 AS carriertype,
            b.Matricule AS personnelnr,
            b.VendorCode,
            1 AS blocked,
            0 AS disabled
        FROM #ToBlock AS b;

        SET @BlocksQueued = @@ROWCOUNT;

        /* 3) MAJ SNAPSHOT (TAMPON) */
        TRUNCATE TABLE dbo.PJ_PRESTATAIRES_IMPORT_TAMPON;

        INSERT INTO dbo.PJ_PRESTATAIRES_IMPORT_TAMPON (Matricule, VendorCode, LastSeen)
        SELECT DISTINCT
            s.personnelnr AS Matricule,
            NULLIF(LTRIM(RTRIM(s.vendor_code)), '') AS VendorCode,
            SYSDATETIME() AS LastSeen
        FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE AS s
        WHERE s.personnelnr IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.vendor_code)), '') IS NOT NULL;

        /* 4) AUDIT create/update */
        INSERT INTO dbo.PJ_PRESTATAIRES_IMPORT_AUDIT (
            RunId, RunTs, ActionType, PersonnelNr, LastName, FirstName, VendorCode
        )
        SELECT
            @RunId,
            @RunTs,
            CASE
                WHEN t.Matricule IS NULL THEN N'CREATE'
                ELSE N'UPDATE'
            END,
            s.personnelnr,
            s.lastname,
            COALESCE(NULLIF(LTRIM(RTRIM(s.firstname)), ''), NULLIF(LTRIM(RTRIM(s.initials)), '')),
            NULLIF(LTRIM(RTRIM(s.vendor_code)), '')
        FROM dbo.PJ_PRESTATAIRES_IMPORT_STAGE s
        LEFT JOIN dbo.PJ_PRESTATAIRES_IMPORT_TAMPON t
               ON t.Matricule = s.personnelnr
              AND t.VendorCode = NULLIF(LTRIM(RTRIM(s.vendor_code)), '')
        WHERE s.personnelnr IS NOT NULL
          AND s.lastname IS NOT NULL
          AND NULLIF(LTRIM(RTRIM(s.vendor_code)), '') IS NOT NULL
          AND s.arrivaldate IS NOT NULL
          AND s.leavedate IS NOT NULL;

        /* 5) HIST */
        INSERT INTO dbo.PJ_PRESTATAIRES_IMPORT_HIST (
            RunId, RunTs, RowsInStage, PersonsQueued, BlocksQueued, SourceFile, BatchId, Notes
        )
        VALUES (
            @RunId, @RunTs, @RowsInStage, @PersonsQueued, @BlocksQueued,
            @SourceFile, @BatchId,
            CONCAT(N'Rejected: ', CONVERT(NVARCHAR(10), (@RowsInStage - @PersonsQueued)))
        );

        COMMIT TRAN;

        SELECT @RunId AS RunId, @RunTs AS RunTs;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRAN;
        THROW;
    END CATCH
END

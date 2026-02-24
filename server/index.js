const express = require('express');
const cors = require('cors');
const sql = require('mssql');

const app = express();
const port = Number(process.env.PORT || 3001);

app.use(cors());
app.use(express.json({ limit: '5mb' }));

function isEmptyValue(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function normalizeSqlParameterName(value) {
  return String(value || '').replace(/^@/, '');
}

function sanitizeProcedureName(input) {
  if (!input || typeof input !== 'string') {
    throw new Error('Procedure name is required.');
  }

  if (!/^[A-Za-z0-9_\.\[\]]+$/.test(input)) {
    throw new Error('Procedure name contains invalid characters.');
  }

  return input;
}

function buildSqlConfig(body) {
  if (!body.server || !body.username || !body.password) {
    throw new Error('Server, username, and password are required.');
  }

  return {
    server: body.server,
    user: body.username,
    password: body.password,
    database: body.database || undefined,
    options: {
      encrypt: false,
      trustServerCertificate: true
    },
    pool: {
      max: 5,
      min: 0,
      idleTimeoutMillis: 30000
    }
  };
}

async function withConnection(config, action) {
  const pool = new sql.ConnectionPool(config);
  await pool.connect();

  try {
    return await action(pool);
  } finally {
    await pool.close();
  }
}

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, message: 'SP connector API is running' });
});

app.post('/api/sp/validate', async (req, res) => {
  const body = req.body || {};

  if (body.spMode && body.spMode !== 'stored_procedure') {
    return res.status(400).json({
      ok: false,
      message: 'This API currently supports SQL Stored Procedure mode only.'
    });
  }

  let procedureName;
  let config;

  try {
    procedureName = sanitizeProcedureName(body.procedureOrList);
    config = buildSqlConfig(body);
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message });
  }

  try {
    const result = await withConnection(config, async (pool) => {
      const request = pool.request();
      const cleanName = procedureName.replace(/[\[\]]/g, '');
      const segments = cleanName.split('.').filter(Boolean);
      const procName = segments[segments.length - 1];
      const schemaName = segments.length > 1 ? segments[segments.length - 2] : null;

      request.input('procName', sql.NVarChar, procName);
      request.input('schemaName', sql.NVarChar, schemaName);

      const existsResult = await request.query(`
        SELECT TOP 1 p.object_id, p.name, s.name AS schema_name
        FROM sys.procedures p
        INNER JOIN sys.schemas s ON p.schema_id = s.schema_id
        WHERE p.name = @procName
          AND (@schemaName IS NULL OR s.name = @schemaName)
      `);

      if (!existsResult.recordset.length) {
        return { exists: false, parameters: [] };
      }

      const objectId = existsResult.recordset[0].object_id;
      const metadataRequest = pool.request();
      metadataRequest.input('objectId', sql.Int, objectId);

      const metadataResult = await metadataRequest.query(`
        SELECT
          prm.name AS parameter_name,
          TYPE_NAME(prm.user_type_id) AS parameter_type,
          prm.max_length AS max_length,
          prm.is_output AS is_output
        FROM sys.parameters prm
        WHERE prm.object_id = @objectId
        ORDER BY prm.parameter_id
      `);

      return {
        exists: true,
        schemaName: existsResult.recordset[0].schema_name,
        procedureName: existsResult.recordset[0].name,
        parameters: metadataResult.recordset.map((row) => ({
          name: row.parameter_name,
          type: row.parameter_type,
          maxLength: row.max_length,
          isOutput: Boolean(row.is_output)
        }))
      };
    });

    if (!result.exists) {
      return res.status(404).json({
        ok: false,
        message: `Connected, but procedure '${procedureName}' was not found.`
      });
    }

    return res.json({
      ok: true,
      message: 'Connection successful and procedure found.',
      procedure: {
        schema: result.schemaName,
        name: result.procedureName
      },
      parameters: result.parameters
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: error.message || 'Validation failed.' });
  }
});

app.post('/api/sp/execute', async (req, res) => {
  const body = req.body || {};

  if (body.spMode && body.spMode !== 'stored_procedure') {
    return res.status(400).json({
      ok: false,
      message: 'Execute endpoint currently supports SQL Stored Procedure mode only.'
    });
  }

  let procedureName;
  let config;

  try {
    procedureName = sanitizeProcedureName(body.procedureOrList);
    config = buildSqlConfig(body);
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message });
  }

  const insertOneByOne = Boolean(body.insertOneByOne);
  const addMissingData = Boolean(body.addMissingData);
  const continueOnError = Boolean(body.continueOnError);
  const rows = Array.isArray(body.rows) ? body.rows : [];
  const parameterMap = body.parameterMap && typeof body.parameterMap === 'object' ? body.parameterMap : {};
  const parameterDefaults =
    body.parameterDefaults && typeof body.parameterDefaults === 'object' ? body.parameterDefaults : {};

  try {
    const result = await withConnection(config, async (pool) => {
      if (!rows.length) {
        await pool.request().execute(procedureName);
        return {
          processed: 1,
          inserted: 1,
          skipped: 0,
          failed: 0,
          details: [{ index: 0, status: 'executed_once' }]
        };
      }

      const details = [];
      let inserted = 0;
      let skipped = 0;
      let failed = 0;

      for (let index = 0; index < rows.length; index++) {
        const row = rows[index];

        try {
          if (!insertOneByOne) {
            throw new Error('Batch mode is not implemented yet. Use insert one by one.');
          }

          const request = pool.request();

          Object.entries(parameterMap).forEach(([sqlParam, sourceField]) => {
            const paramName = normalizeSqlParameterName(sqlParam);
            const sourceName = String(sourceField || '');
            const sourceValue = sourceName ? row[sourceName] : undefined;
            const hasDefault = Object.prototype.hasOwnProperty.call(parameterDefaults, sqlParam);
            const resolvedValue = isEmptyValue(sourceValue) && hasDefault ? parameterDefaults[sqlParam] : sourceValue;

            request.input(paramName, resolvedValue);
          });

          request.input('addMissingData', addMissingData);

          await request.execute(procedureName);
          inserted += 1;
          details.push({ index, status: 'inserted' });
        } catch (error) {
          failed += 1;
          details.push({ index, status: 'failed', error: error.message || 'Unknown row error' });

          if (!continueOnError) {
            break;
          }
        }
      }

      skipped = Math.max(rows.length - inserted - failed, 0);

      return {
        processed: rows.length,
        inserted,
        skipped,
        failed,
        details
      };
    });

    return res.json({ ok: true, message: 'Execution completed.', result });
  } catch (error) {
    return res.status(500).json({ ok: false, message: error.message || 'Execution failed.' });
  }
});

app.listen(port, () => {
  console.log(`SP connector API listening on http://localhost:${port}`);
});

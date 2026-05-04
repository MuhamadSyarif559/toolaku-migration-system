const express = require('express');
const cors = require('cors');
const sql = require('mssql');
const fs = require('fs/promises');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const app = express();
const port = Number(process.env.PORT || 3001);

app.use(cors());
app.use(express.json({ limit: '100mb' }));

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

function ensureArray(value) {
  return Array.isArray(value) ? value : [];
}

function sanitizeFileName(fileName, fallback) {
  const base = path.basename(String(fileName || fallback));
  return base.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_');
}

const DATA_CLEANING_TOOLS = [
  {
    mode: 'grn_header_refs',
    category: 'GRN',
    title: 'GRN Header References',
    description: 'Fill supplier, term, location, and currency IDs in the GRN header workbook.',
    script: 'excel_fill_grn_refs.py',
    outputName: 'grn-header-with-reference-ids.xlsx',
    requiredFields: ['source', 'supplier', 'term', 'location', 'currency'],
    inputs: [
      { key: 'source', label: 'GRN Header Workbook', hint: 'Main GRN header migration workbook' },
      { key: 'supplier', label: 'Supplier Reference', hint: 'Supplier list.xlsx' },
      { key: 'term', label: 'Term Reference', hint: 'Term list.xlsx' },
      { key: 'location', label: 'Location Reference', hint: 'list of location.xlsx' },
      { key: 'currency', label: 'Currency Reference', hint: 'Currency List.xlsx' }
    ],
    buildArgs(savedFiles, outputPath) {
      return [
        '--source', savedFiles.source,
        '--supplier', savedFiles.supplier,
        '--term', savedFiles.term,
        '--location', savedFiles.location,
        '--currency', savedFiles.currency,
        '--output', outputPath
      ];
    }
  },
  {
    mode: 'grn_detail_refs',
    category: 'GRN',
    title: 'GRN Detail References',
    description: 'Fill stock, tax, location, and UOM IDs in the GRN detail workbook.',
    script: 'excel_fill_grn_detail_refs.py',
    outputName: 'grn-detail-with-reference-data-and-uomid.xlsx',
    requiredFields: ['source', 'stock', 'tax', 'location', 'uom'],
    inputs: [
      { key: 'source', label: 'GRN Detail Workbook', hint: 'Main GRN details migration workbook' },
      { key: 'stock', label: 'Stock Reference', hint: 'StockRefExcel.xlsx' },
      { key: 'tax', label: 'Tax Reference', hint: 'taxIdwithCode.xlsx' },
      { key: 'location', label: 'Location Reference', hint: 'list of location.xlsx' },
      { key: 'uom', label: 'UOM Reference', hint: 'alluomdetails.xlsx or equivalent UOM master' }
    ],
    buildArgs(savedFiles, outputPath) {
      return [
        '--source', savedFiles.source,
        '--stock', savedFiles.stock,
        '--tax', savedFiles.tax,
        '--location', savedFiles.location,
        '--uom', savedFiles.uom,
        '--output', outputPath
      ];
    }
  },
  {
    mode: 'grn_detail_cleanup',
    category: 'GRN',
    title: 'GRN Detail Cleanup',
    description: 'Merge note descriptions and remove zero-quantity rows without UOM, while preserving location rows.',
    script: 'cleanup_grn_detail_rows.py',
    outputName: 'grn-detail-cleaned.xlsx',
    requiredFields: ['source'],
    inputs: [
      { key: 'source', label: 'Cleanable Detail Workbook', hint: 'Workbook produced by the detail reference step' }
    ],
    buildArgs(savedFiles, outputPath) {
      return ['--source', savedFiles.source, '--output', outputPath];
    }
  }
];

function getDataCleaningConfig(mode, tempDir, files, requestedOutputName) {
  const rootDir = path.resolve(__dirname, '..', '..', '..');
  const config = DATA_CLEANING_TOOLS.find((item) => item.mode === mode);
  if (!config) {
    throw new Error('Unsupported data cleaning mode.');
  }

  const savedFiles = {};
  for (const file of files) {
    if (!file.fieldName || !file.fileName || !file.contentBase64) {
      throw new Error('Each uploaded file must include fieldName, fileName, and contentBase64.');
    }
    const safeFileName = sanitizeFileName(file.fileName, `${file.fieldName}.xlsx`);
    savedFiles[file.fieldName] = path.join(tempDir, safeFileName);
  }

  for (const field of config.requiredFields) {
    if (!savedFiles[field]) {
      throw new Error(`Missing required file: ${field}`);
    }
  }

  return {
    scriptPath: path.join(rootDir, config.script),
    outputPath: path.join(tempDir, sanitizeFileName(requestedOutputName || config.outputName, 'output.xlsx')),
    downloadFileName: sanitizeFileName(requestedOutputName || config.outputName, 'output.xlsx'),
    args: config.buildArgs(
      savedFiles,
      path.join(tempDir, sanitizeFileName(requestedOutputName || config.outputName, 'output.xlsx'))
    )
  };
}

async function runPythonScript(scriptPath, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.env.PYTHON_PATH || 'python', [scriptPath, ...args], {
      windowsHide: true
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString();
    });

    child.on('error', (error) => {
      reject(error);
    });

    child.on('close', (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }

      reject(new Error(stderr.trim() || stdout.trim() || `Python script failed with exit code ${code}.`));
    });
  });
}

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, message: 'SP connector API is running' });
});

app.get('/api/data-cleaning/catalog', (_req, res) => {
  res.json({
    ok: true,
    tools: DATA_CLEANING_TOOLS.map((tool) => ({
      mode: tool.mode,
      category: tool.category,
      title: tool.title,
      description: tool.description,
      outputName: tool.outputName,
      requiredFields: tool.requiredFields,
      inputs: tool.inputs
    }))
  });
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

app.post('/api/data-cleaning/run', async (req, res) => {
  const body = req.body || {};
  const files = ensureArray(body.files);
  const mode = String(body.mode || '').trim();
  const requestedOutputName = String(body.outputFileName || '').trim();

  if (!mode) {
    return res.status(400).json({ ok: false, message: 'Cleaning mode is required.' });
  }

  if (!files.length) {
    return res.status(400).json({ ok: false, message: 'At least one uploaded file is required.' });
  }

  let tempDir = '';

  try {
    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'toolaku-clean-'));

    for (const file of files) {
      const outputPath = path.join(tempDir, sanitizeFileName(file.fileName, `${file.fieldName}.xlsx`));
      await fs.writeFile(outputPath, Buffer.from(String(file.contentBase64), 'base64'));
    }

    const config = getDataCleaningConfig(mode, tempDir, files, requestedOutputName);
    const { stdout, stderr } = await runPythonScript(config.scriptPath, config.args);
    const outputBuffer = await fs.readFile(config.outputPath);

    return res.json({
      ok: true,
      message: 'Data cleaning completed.',
      mode,
      downloadFileName: config.downloadFileName,
      outputBase64: outputBuffer.toString('base64'),
      stdout,
      stderr
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      message: error instanceof Error ? error.message : 'Data cleaning failed.'
    });
  } finally {
    if (tempDir) {
      await fs.rmdir(tempDir, { recursive: true }).catch(() => undefined);
    }
  }
});

app.listen(port, () => {
  console.log(`SP connector API listening on http://localhost:${port}`);
});

import path from 'node:path';
import { finalizePresentation } from '/root/.codex/skills/builtins/presentations/container_tools/artifact_tool_utils.mjs';
const root='/workspace/scratch/242dd116dead/gaza';
const skill='/root/.codex/skills/builtins/presentations';
const result=await finalizePresentation({
workspaceDir:root,candidatePath:path.join(root,'build/candidate.pptx'),finalPath:path.join(root,'delivery/akuf_gazastreifen_praesentation.pptx'),
pythonExecutable:process.env.CODEX_PRIMARY_RUNTIME_PYTHON,
integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),
layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
layoutArgs:['--expected-slide-size-emu','12191695,6858000','--validate-bullet-geometry','--validate-heading-fit'],
explicitTotalSlideCount:12,requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[],
fontPolicy:{basis:'design',families:['DejaVu Sans']},verifyArtifactToolImport:true,
receiptPath:path.join(root,'build/validation-final.json')
});
console.log(JSON.stringify(result));

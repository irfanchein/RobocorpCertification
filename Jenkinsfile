pipeline {
    agent {
        kubernetes {
            label 'rpa_gl_r2r_lease_reconciliation'
            yaml '''
apiVersion: v1
kind: Pod
spec:
    serviceAccountName: jenkins
    containers:
    - name: "jnlp"
      resources:
        requests:
            cpu: 200m
            memory: 400Mi
        limits:
            cpu: 400m
            memory: 800Mi
    - name: awscli
      image: amazon/aws-cli:latest
      imagePullPolicy: IfNotPresent
      resources:
        requests:
            cpu: 200m
            memory: 400Mi
        limits:
            cpu: 400m
            memory: 800Mi
      command:
        - cat
      tty: true
'''
        }
    }
    environment {
        ENVIRONMENT = getEnv(env.BRANCH_NAME)
        WORKSPACE_ID   = getWorkspaceId(env.BRANCH_NAME)
        ROBOT_ID = getRobotId(env.BRANCH_NAME)
        API_KEY_ROBOT = getAPIKey(env.BRANCH_NAME)
        CREDENTIAL_ROBOT = getCredential(env.BRANCH_NAME)
    }
    stages {
        stage('Build') {
            steps {
                container('awscli'){
                    sh'''set +x
                    | echo "Start Build"
                    | curl -o rcc https://downloads.robocorp.com/rcc/releases/latest/linux64/rcc
                    | chmod a+x rcc
                    | mv rcc /usr/local/bin/
                    '''.stripMargin()

                    sh(script: "rcc configure credentials '${API_KEY_ROBOT}'")

                    sh(script: "rcc cloud push -r ${ROBOT_ID} -w ${WORKSPACE_ID} -a ${CREDENTIAL_ROBOT}:${API_URL}")

                    echo "Upload to control Room Done"
                }
            }
        }  
    }
}

def getEnv(branchName) {
    if(branchName.equals("dev")) return "dev";
	if(branchName.equals("qa")) return "qas";
	if(branchName.equals("prod")) return "prd";
	return "";
}

def getWorkspaceId(branchName) {
    if(branchName.equals("dev")) return env["WORKSPACE_ID_DEV"];
	if(branchName.equals("qa")) return env["WORKSPACE_ID_QA"];
	if(branchName.equals("prod")) return env["WORKSPACE_ID_PROD"];
	return "";
}

def getRobotId(branchName) {
    if(branchName.equals("dev")) return env["ROBOT_ID_DEV"];
	if(branchName.equals("qa")) return env["ROBOT_ID_QA"];
	if(branchName.equals("prod")) return env["ROBOT_ID_PROD"];
	return "";
}

def getAPIKey(branchName) {
    if(branchName.equals("dev")) return env["API_KEY_DEV"];
	if(branchName.equals("qa")) return env["API_KEY_QA"];
	if(branchName.equals("prod")) return env["API_KEY_PROD"];
	return "";
}

def getCredential(branchName) {
    if(branchName.equals("dev")) return env["CREDENTIAL_DEV"];
	if(branchName.equals("qa")) return env["CREDENTIAL_QA"];
	if(branchName.equals("prod")) return env["CREDENTIAL_PROD"];
	return "";
}
